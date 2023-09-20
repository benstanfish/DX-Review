Attribute VB_Name = "dxreview3"
Const module_version = "2.0.0"

' Module requires references to:
'    Microsoft XML, v6.0 (msxml6.dll) - XML parsing functions
'    Microsoft Scripting Runtime (scrrun.dll) - Dictionaries
'    Microsoft VBScripting Regular Expressions 5.5 (vbscript.dll)
'    Microsoft Visual Basic for Applications Extensibility 5.3


' 1. Loading and verifying XML files

Public Function ParseXML(file_path As String) As IXMLDOMElement
    ' Load an XML file and return the root node as IXMLDOMElement, else return NOTHING
    Dim xml_doc As DOMDocument60
    Dim temp_element As IXMLDOMElement
    Set xml_doc = New DOMDocument60
    With xml_doc
        .validateOnParse = False
        If .Load(file_path) = False Then
            Set ParseXML = Nothing
        Else
            .Load file_path
            Set temp_element = .DocumentElement
            If VerifyRoot(temp_element) Then Set ParseXML = temp_element
        End If
    End With
End Function

Public Function VerifyRoot(ByVal root As IXMLDOMElement) As Boolean
    ' Return TRUE if the XML file is a Dr Checks/ProjNet document
    If root Is Nothing Then
        VerifyRoot = False
    ElseIf root.nodeName = "ProjNet" Then
        VerifyRoot = True
    End If
End Function

Public Function GetFilePath() As String
    ' Returns EMPTY if user cancels, otherwise returns path string
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "XML", "*.xml?"
        .Title = "Choose an XML file"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetFilePath = .SelectedItems(1)
    End With
End Function

Public Function GetFolderPath() As String
    ' Returns EMPTY if user cancels, otherwise returns path string
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetFolderPath = .SelectedItems(1)
    End With
End Function

' 2. Creating Workbooks, renaming workesheets, etc.

Function CreateWorkbook(save_path As String, Optional workbook_name As String = "DrChecks Summary Report", Optional include_timestamp As Boolean = True) As Workbook
    ' Return a  new Workbook object with the provided name and appends with a timestamp as noted.
    Dim combined_workbook As Workbook
    Dim file_name As String
    Set combined_workbook = Workbooks.Add
    Application.DisplayAlerts = False
    With combined_workbook
        .Title = workbook_name
        If include_timestamp Then
            file_name = save_path & "\" & workbook_name & " " & Format(Now(), "YYYY-MM-DD hh-mm-ss") & ".xlsx"
        Else
            file_name = save_path & "\" & workbook_name & ".xlsx"
        End If
        .SaveAs Filename:=file_name, FileFormat:=xlOpenXMLWorkbook
    End With
    Application.DisplayAlerts = True
    Set CreateWorkbook = combined_workbook
End Function

Sub RenameSheet(ByVal target_sheet As Worksheet, ByVal root_element As IXMLDOMElement)
    ' Renames a Worksheet (tab) with the value of the <ReviewName> node of an XML file.
    ' It limits the name to the maximum permitted character count (31 - 4 = 27) and removes
    ' illegal characters from the name
    
    Dim illegal_characters As Variant
    Dim new_sheet_name As String
    Dim i As Long
    'Create array of characters that are not permitted in worksheet names
    illegal_characters = Array("/", "\", "?", "*", ":", "[", "]")
    new_sheet_name = root_element.SelectSingleNode("DrChecks/ReviewName").Text
    If Len(new_sheet_name) > 27 Then new_sheet_name = Left(new_sheet_name, 27)
    For i = LBound(illegal_characters) To UBound(illegal_characters)
        new_sheet_name = Replace(new_sheet_name, illegal_characters(i), "")
    Next
    On Error GoTo dump
    target_sheet.Name = new_sheet_name
dump:
End Sub

Public Function GetRootFromXML(ByVal file_path As String) As IXMLDOMElement
    ' Load an XML file and return the root node as IXMLDOMElement, else return Nothing
    Dim xml_doc As DOMDocument60
    Set xml_doc = New DOMDocument60
    With xml_doc
        .validateOnParse = False
        If .Load(file_path) = False Then
            Set GetRootFromXML = Nothing
        Else
            .Load file_path
            Set GetRootFromXML = .DocumentElement
        End If
    End With
End Function

Public Function BuildFromXML(ByVal file_path As String) As IXMLDOMElement
    Dim root As IXMLDOMElement
    Dim projInfo As New ProjectInfo
    Dim all_comments As New Comments
    Dim user_notes As New UserNotes

    Set root = GetRootFromXML(file_path)
    If VerifyRoot(root) Then
        ActiveSheet.Cells.Clear
    
        projInfo.CreateFromNode root.SelectSingleNode("DrChecks")
        projInfo.PasteData ActiveSheet.Range("E1")
    
        all_comments.CreateFromRootElement root
        all_comments.PasteData ActiveSheet.Range("E7")
    
        user_notes.PasteData ActiveSheet.Range("A7"), all_comments.Count
    
        all_comments.ApplyFormats
        user_notes.ApplyFormats

        Set BuildFromXML = root
    End If
End Function

Sub ImportFile()
    Application.ScreenUpdating = False

    Dim fso As New FileSystemObject
    Dim file_path As String
    Dim folder_path As String
    Dim wb As Workbook
    
    file_path = GetFilePath
    If file_path <> Empty Then
        Set root = GetRootFromXML(file_path)
        If VerifyRoot(root) Then
            folder_path = fso.GetParentFolderName(file_path)
            Set wb = CreateWorkbook(folder_path)
            wb.Sheets.Add After:=wb.Sheets(wb.Sheets.Count)
            Set current_sheet = wb.Sheets(wb.Sheets.Count)
            Set root = BuildFromXML(file_path)
            RenameSheet current_sheet, root
            WriteDevInfo wb
        End If
    End If

    Application.ScreenUpdating = True

End Sub

Sub ImportMultipleFiles()
    Application.ScreenUpdating = False
    
    
    Dim fso As FileSystemObject
    Dim a_folder As Folder
    Dim a_file As File
    Dim file_path As String
    Dim folder_path As String

    Dim wb As Workbook
    Dim safe_files As Collection
    
    Set fso = New FileSystemObject
       
    folder_path = GetFolderPath
    If folder_path <> Empty Then
        Set wb = CreateWorkbook(folder_path)
        For Each a_file In fso.GetFolder(folder_path).Files
            Set root = GetRootFromXML(a_file)
            If VerifyRoot(root) = True Then
                folder_path = fso.GetParentFolderName(a_file)
                
                wb.Sheets.Add After:=wb.Sheets(wb.Sheets.Count)
                Set current_sheet = wb.Sheets(wb.Sheets.Count)
                Set root = BuildFromXML(a_file)
                RenameSheet current_sheet, root
            End If
        Next
        WriteDevInfo wb
    End If

    Application.ScreenUpdating = True
    
End Sub

' 3. Finishing

Sub FindInExplorer(ByVal folder_path As String, Optional is_in_focus As Boolean = False)
    If folder_path <> "" Then
        If is_in_focus Then
            Shell "C:\WINDOWS\explorer.exe """ & folder_path & "", vbNormalFocus
        Else
            Shell "C:\WINDOWS\explorer.exe """ & folder_path & "", vbNormalNoFocus
        End If
    End If
End Sub

Sub WriteDevInfo(target_workbook As Workbook)
    Dim header_array, values_array As Variant
    Dim start_cell As Range
    Dim i As Integer
    header_array = Array("Program", "Module Name", "Version", _
                        "Author", "Email", "Github", "License", "References", , "Run Date")

    values_array = Array("DX Review", "dxreview2", module_version, _
                        "Ben Fisher", "benstanfish@gmail.com", "https://github.com/benstanfish/DX-Review", _
                        "GNU General Public License v3.0", _
                        "Microsoft XML v6 (msxml.dll), Microsoft Scripting Runtime (scrrun.dll)", , CDate(Now))
                        
    With target_workbook.Sheets(1)
        .Cells.Delete Shift:=xlUp
        Set start_cell = .Range("A1")
        For i = LBound(header_array, 1) To UBound(header_array, 1)
            start_cell.Offset(i, 0) = header_array(i)
            start_cell.Offset(i, 1) = values_array(i)
        Next
        With start_cell.Columns(1).EntireColumn
            .Font.Bold = True
            .AutoFit
        End With
        start_cell.Offset(0, 1).EntireColumn.ColumnWidth = 15
        .Cells.HorizontalAlignment = xlHAlignLeft
        .Name = "DevInfo"
        If target_workbook.Sheets.Count > 1 Then
            .Visible = xlSheetVeryHidden
        End If
    End With
End Sub

' 4. Miscellaneous

Private Sub ExportModules()

    Dim me_path As String
    Dim comp As VBIDE.VBComponent
    
    me_path = Application.ActiveWorkbook.Path & "\"
    
    For Each comp In ActiveWorkbook.VBProject.VBComponents
        is_export = True
        Select Case comp.Type
            Case vbext_ct_ClassModule
                comp.Export me_path & comp.Name & ".cls"
            Case vbext_ct_MSForm
                comp.Export me_path & comp.Name & ".frm"
            Case vbext_ct_StdModule
                comp.Export me_path & comp.Name & ".bas"
            Case vbext_ct_Document
                ' Don't export
        End Select
    Next
End Sub














