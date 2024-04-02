Attribute VB_Name = "dxreview"
Public Const mod_name As String = "DxReview"
Public Const module_author As String = "Ben Fisher"
Public Const module_email As String = "benstanfish@hotmail.com"
Public Const module_version As String = "4.6.5"
Public Const module_date As Date = #4/2/2024#
Public Const module_dependencies = "Microsoft XML, v6.0 (msxml6.dll) - XML parsing functions" & vbCrLf & _
                                    "Microsoft Scripting Runtime (scrrun.dll) - Dictionaries" & vbCrLf & _
                                    "Microsoft VBScript Regular Expressions 5.5 (vbscript.dll)" & vbCrLf & _
                                    "Microsoft Visual Basic for Applications Extensibility 5.3" & vbCrLf & _
                                    "Microsoft HTML Object Library (MSHTML.tlb)"

Public Const PROJECTINFOTARGETCELL As String = "H1"
Public Const ALLCOMMENTSTARGETCELL  As String = "H11"
Public Const USERNOTESTARGETCELL  As String = "A11"

Public Const SOLIDTRIANGLEUP As Long = 9650
Public Const TRIANGLEUP As Long = 9651
Public Const SOLIDTRIANGLERIGHT As Long = 9658
Public Const TRIANGLERIGHT As Long = 9655
Public Const SOLIDTRIANGLEDOWN As Long = 9660
Public Const TRIANGLEDOWN As Long = 9661
Public Const SOLIDTRIANGLELEFT As Long = 9664
Public Const TRIANGLELEFT As Long = 9665
Public Const DAIMARU As Long = 9898
Public Const SOLIDDAIMARU As Long = 9899


Private Sub UpdateVersionNumber()

    With ThisWorkbook.Sheets("Macros")
        .Unprotect Password:=""
        With .Range("K3")
            .Value = mod_name & " v" & module_version
            .HorizontalAlignment = xlHAlignRight
            .Font.Size = 9
            .Font.Italic = True
        End With
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, UserInterfaceOnly:=True, Password:=""
        .EnableSelection = xlUnlockedCells
        .EnableOutlining = True
    End With

End Sub

Private Sub UnprotectSheet()
    ThisWorkbook.Sheets("Macros").Unprotect Password:=""
End Sub

Public Sub GroupedColumnTriangles(Optional dummy As Long = 0)

    With Range(dxreview.USERNOTESTARGETCELL).Offset(-1, 0)
        .Value = ChrW(TRIANGLERIGHT)
        .HorizontalAlignment = xlHAlignRight
    End With
    Range(dxreview.ALLCOMMENTSTARGETCELL).Offset(-1, 0).Value = ChrW(TRIANGLELEFT)

    With Range(dxreview.ALLCOMMENTSTARGETCELL).Offset(-1, 4)
        .Value = ChrW(TRIANGLERIGHT)
        .HorizontalAlignment = xlHAlignRight
    End With
    Range(dxreview.ALLCOMMENTSTARGETCELL).Offset(-1, 10).Value = ChrW(TRIANGLELEFT)

End Sub


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

Public Function GetXMLPath() As String
    ' Returns EMPTY if user cancels, otherwise returns path string
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "XML", "*.xml?"
        .Title = "Choose an XML file"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetXMLPath = .SelectedItems(1)
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

Function ReadFileDate(aPath As String)
    Dim fso As FileSystemObject
    Dim aFile As File
    If aPath <> "" Then
        Set fso = New FileSystemObject
        Set aFile = fso.GetFile(aPath)
        ReadFileDate = aFile.DateCreated
        Set fso = Nothing
        Set aFile = Nothing
    Else
        ReadFileDate = Now()
    End If
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
    target_sheet.Name = new_sheet_name & IterateSheetName(new_sheet_name)
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
    
        projInfo.CreateFromNode root.SelectSingleNode("DrChecks"), file_path
        projInfo.PasteData ActiveSheet.Range(PROJECTINFOTARGETCELL)
    
        all_comments.CreateFromRootElement root
        all_comments.PasteData ActiveSheet.Range(ALLCOMMENTSTARGETCELL)
    
        user_notes.PasteData ActiveSheet.Range(USERNOTESTARGETCELL), all_comments.Count
    
        all_comments.ApplyFormats
        user_notes.ApplyFormats

        Dim aTable As ListObject
        Set aTable = ActiveSheet.ListObjects.Add(SourceType:=xlSrcRange, _
            Source:=ActiveSheet.Range(Range(USERNOTESTARGETCELL), _
                Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count - 1)), _
            XlListObjectHasHeaders:=xlYes)
        
        aTable.Name = IterateTableName("Comments")
        aTable.TableStyle = ""

        InsertDropdown aTable, "Proposed Status", "Concur, Non-concur, For Information Only, Check and Resolve"
        InsertDropdown aTable, "State", "Working, Ready, Done, NA"
        ApplyConditionalFormats aTable, "State", "Working, Ready, Done, NA"

        Dim refRng As Range
        With ActiveSheet.ListObjects(1)
            Set refRng = Union(.ListColumns("Source").Range, .ListColumns("Reference").Range, .ListColumns("Sheet").Range, .ListColumns("Spec").Range, .ListColumns("Section").Range)
        End With
        With refRng
            .Interior.Color = webcolors.LEMONCHIFFON
            .Font.Color = webcolors.SADDLEBROWN
        End With

        Set BuildFromXML = root
    End If
End Function

Sub ImportFile()
    Application.ScreenUpdating = False

    Dim fso As New FileSystemObject
    Dim file_path As String
    Dim folder_path As String
    Dim wb As Workbook
    
    file_path = GetXMLPath
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
            GenerateNewStatSheets wb
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
        GenerateNewStatSheets wb
    End If

    Application.ScreenUpdating = True
    
End Sub

'NOTE: IterateSheetName() moved below

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

    values_array = Array("DX Review", mod_name, module_version, _
                        module_author, module_email, "https://github.com/benstanfish/DX-Review", _
                        "GNU General Public License v3.0", _
                        module_dependencies, , CDate(Now))
                        
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


'==============================  HELPER METHODS  ===============================

Function ParseToArray(namedConstant As String) As Variant
    ParseToArray = Split(namedConstant, ", ")
End Function

Function ParseToLongArray(namedConstant As String) As Variant
    Dim arr As Variant
    Dim arr2 As Variant
    Dim i As Long
    
    arr = ParseToArray(namedConstant)
    ReDim arr2(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        arr2(i) = CLng(arr(i))
    Next

    ParseToLongArray = arr2
End Function

'===================  VALIDATION AND CONDITIONAL FORMATTING  ===================

Sub InsertDropdown(aTable As ListObject, _
                    Optional targetColumn As String = "State", _
                    Optional selectionSet As String = "Ongoing, Ready, Done", _
                    Optional suppressError As Boolean = False)
    'NOTE: This method inserts a validation list with values parsed from the selectionSet, into
    ' the cells of the targetColumn as dropdown lists. USE: combine with conditional formatting.
    ' selectionSet must be a single string with options seperated by a comma and space.
            
    ' Test for empty table
    Dim temporaryRow As Boolean
    If aTable.DataBodyRange Is Nothing Then
        aTable.ListRows.Add
        temporaryRow = True
    End If
    
    With aTable.ListColumns(targetColumn).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=selectionSet
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Disallowed Input"
        .InputMessage = ""
        .ErrorMessage = "Please select from the options: " & selectionSet
        .ShowInput = True
        .ShowError = suppressError
    End With
    
    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete
 
End Sub

Sub ApplyConditionalFormats(aTable As ListObject, _
                        Optional targetColumn As String = "State", _
                        Optional selectionSet As String = "Ongoing, Ready, Done", _
                        Optional hasSecondaryFormats As Boolean = True)
    'NOTE: This method inserts conditional formatting based on the validation "dropdown" list
    ' values in the targetColumn. Values must match those in selectionSet (not case sensitive).
    ' hasSecondaryFormats highlights the whole row with accent scheme, while the main
    ' condition only highlights the targetColumn values. The user must manually update
    ' preferences here in this method.
    
    
    ' Test for empty table
    Dim temporaryRow As Boolean
    If aTable.DataBodyRange Is Nothing Then
        aTable.ListRows.Add
        temporaryRow = True
    End If
    
    Dim choices As Variant
    Dim i As Long

    choices = ParseToArray(selectionSet)
    For i = LBound(choices) To UBound(choices)
        choices(i) = """" & choices(i) & """"
    Next

    Dim statusColumn As Range
    Set statusColumn = aTable.ListColumns(targetColumn).DataBodyRange


    Dim firstCell As String
    firstCell = "$" & Replace(statusColumn(1).Address, "$", "")

    '---------------  MAIN CONDITIONS  ----------------
    statusColumn.FormatConditions.Delete

    For i = LBound(choices) To UBound(choices)
        statusColumn.FormatConditions.Add Type:=xlExpression, _
            Formula1:="=IF(LOWER(" & firstCell & ")=" & choices(i) & ",TRUE,FALSE)"
    Next

    With statusColumn.FormatConditions(1)
        .Interior.Color = webcolors.DANGER
        .Font.Color = ContrastText(.Interior.Color, webcolors.DANGER_DARKER, webcolors.DANGER)
        .Font.Bold = True
    End With

    With statusColumn.FormatConditions(2)
        .Interior.Color = webcolors.WARNING
        .Font.Color = ContrastText(.Interior.Color, webcolors.WARNING_DARKER, webcolors.WARNING)
        .Font.Bold = True
    End With

    With statusColumn.FormatConditions(3)
        .Interior.Color = webcolors.SUCCESS
        .Font.Color = ContrastText(.Interior.Color, webcolors.SUCCESS_DARKER, webcolors.SUCCESS)
        .Font.Bold = True
    End With
    
    With statusColumn.FormatConditions(4)
        .Interior.Color = webcolors.SECONDARY_LIGHT
        .Font.Color = ContrastText(.Interior.Color, webcolors.SECONDARY_DARKER, webcolors.SECONDARY)
        .Font.Bold = True
    End With

    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete

End Sub


'============================  GENERATIVE METHODS  =============================

Function IterateTableName(baseName As String)

    Dim maxIndex As Long
    Dim aTable As ListObject
    For Each sht In ActiveWorkbook.Sheets
        For Each aTable In sht.ListObjects
            If Left(aTable.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
        Next
    Next
    If maxIndex = 0 Then IterateTableName = baseName & "" Else IterateTableName = baseName & maxIndex

End Function

Function IterateSheetName(baseName As String)

    Dim maxIndex As Long
    For Each sht In ActiveWorkbook.Sheets
        If Left(sht.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
    Next
    If maxIndex = 0 Then IterateSheetName = "" Else IterateSheetName = maxIndex

End Function

Sub CopyWorksheetChangeCode(sht As Worksheet)
    'Requires reference to 'Microsoft Visual Basic for Applications Extensibility 5.3"
    'and you must check YES to "Trust Access to VBA Object Model" in Macro Security Settings
    
    Dim VBAEditor As VBIDE.VBE
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComp2 As VBIDE.VBComponent

    Set VBAEditor = Application.VBE
    Set VBProj = VBAEditor.ActiveVBProject
    Set VBComp = VBProj.VBComponents("Sheet2")  ' This should be the initial "POAM Log" sheet
                                                ' even if renamed or shifted by the user.
    Set VBComp2 = VBProj.VBComponents(sht.CodeName)
    
    codeString = VBComp.CodeModule.Lines(1, VBComp.CodeModule.CountOfLines)
    
    VBComp2.CodeModule.DeleteLines 1, VBComp2.CodeModule.CountOfLines
    VBComp2.CodeModule.InsertLines 1, codeString
        
End Sub






