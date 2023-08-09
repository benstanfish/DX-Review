Attribute VB_Name = "dxreviewv2"

Const version as String = "1.0.2"
Const author as String = "Ben Fisher"

Const target_cell_address = "D1"

Const xxlarge_column = 50
Const xlarge_column = 40
Const large_column = 30
Const medium_column = 20
Const small_column = 10
Const xsmall_column = 5

Const color_aliceblue As Long = 16775408
Const color_honeydew As Long = 15794160
Const color_lemonchiffon As Long = 13499135
Const color_gold As Long = 55295
Const color_gainsboro As Long = 14474460
Const color_silver As Long = 12632256
Const color_whitesmoke As Long = 16119285

'#########################################################################
'
'                           Utility Methods
'
'#########################################################################

Private Sub set_max_RowHeight(a_range as Range, Optional max_row_height as Double = 75)
    For Each a_row In a_range.Rows
        If a_row.RowHeight > max_row_height Then a_row.RowHeight = max_row_height
    Next
End Sub

Function count_all_comments(root_element As IXMLDOMElement) as Long
    'Returns a count of all <comment> elements that are children of the <Comments>
    'node in the XML file. Note: using XPATH="comment" results in duplicates because
    'all evaulation and backcheck nodes include the parent <comment> as a child!!!
    count_all_comments = root_element.selectNodes("Comments/comment").Length
End Function

Function get_max_element_count(element_xpath as String, root_element As IXMLDOMElement)
    'Used to calculate the maximum number of <evaulation*> or <backcheck*> elements for
    'any given <comment> in the XML file. The element_xpath is typically either
    '"Comments/comment/evaluations" or "Comments/comment/backchecks".
    Dim element_selection As IXMLDOMSelection
    Dim i, max_count As Long
    Set element_selection = root_element.selectNodes(element_xpath)
    For i = 0 To element_selection.Length - 1
        If max_count < element_selection(i).ChildNodes.Length Then
            max_count = element_selection(i).ChildNodes.Length
        End If
    Next
    element_selection = max_count
End Function

Function count_days_open(comment_node As IXMLDOMElement) As Long
    'If the comment status is "open" this function returns the number of days
    'since the comment was originally created. If the comment is closed, it
    'returns the number of days between the create date and the date of the
    'last backcheck comment.
    Dim comment_status As String
    Dim comment_created_date, final_backcheck_date As Date
    Dim backchecks_section As IXMLDOMSelection
    Dim final_backcheck As IXMLDOMElement
    Dim backchecks_count As Long
    comment_status = LCase(comment_node.selectSingleNode("status").Text)
    comment_created_date = CDate(comment_node.selectSingleNode("createdOn").Text)
    If comment_status = "closed" Then
        Set backchecks_section = comment_node.selectNodes("backchecks/*")
        backchecks_count = backchecks_section.Length
        Set final_backcheck = backchecks_section.Item(backchecks_count - 1)
        final_backcheck_date = CDate(final_backcheck.selectSingleNode("createdOn").Text)
        count_days_open = DateDiff("d", comment_created_date, final_backcheck_date)
    Else
        count_days_open = DateDiff("d", comment_created_date, Now())
    End If
End Function

Sub rename_sheet(root_element As IXMLDOMElement, a_sheet As Worksheet)
    'Renames a Worksheet (tab) with the value of the <ReviewName> node of an XML file.
    'It limits the name to the maximum permitted character count (31) and removes 
    'illegal characters from the name.
    Dim illegal_characters As Variant
    Dim new_sheet_name as String
    Dim i as Long
    'Create array of characters that are not permitted in worksheet names
    illegal_characters = Array("/", "\", "?", "*", ":", "[", "]")
    new_sheet_name = root_element.selectSingleNode("DrChecks/ReviewName").Text
    If Len(new_sheet_name) > 31 Then new_sheet_name = Left(new_sheet_name, 31)
    For i = LBound(illegal_characters) To UBound(illegal_characters)
        new_sheet_name = Replace(new_sheet_name, illegal_characters(i), "")
    Next
    On Error GoTo dump
    a_sheet.Name = new_sheet_name
dump:
End Sub

Public Function long_to_rgb(color_as_Long As Long) as Variant
    Dim r, b, g as Long
    b = color_as_Long \ 65536
    g = (color_as_Long - b * 65536) \ 256
    r = color_as_Long - b * 65536 - g * 256
    long_to_rgb = Array(r, g, b)
End Function

Public Function rgb_to_long(rgb_arr As Variant) As Long
    rgb_to_long = RGB(rgb_arr(0), rgb_arr(1), rgb_arr(2))
End Function

Public Function select_contrast_font(background_color As Long) as Long
    'Returns the color white or black as a Long, determined as the most appropriate
    'constrasting font color based on the supplied background_color. Based on the
    'W3 visibility recommendations, ref. https://www.w3.org/TR/AERT/#color-contrast
    Dim rgb_arr As Variant
    Dim color_constant As Long
    Dim color_brightness As Double
    rgb_arr = long_to_rgb(background_color)
    color_brightness = (0.299 * rgb_arr(0) + 0.587 * rgb_arr(1) + 0.114 * rgb_arr(2)) / 255
    If color_brightness > 0.55 Then color_constant = vbBlack Else color_constant = vbWhite
    select_contrast_font = color_constant
End Function



'#########################################################################
'
'                             MAIN METHOD
'
'#########################################################################

Sub MAIN()

    perform_on_each_file

End Sub

Function create_workbook(save_path as String, 
                         Optional workbook_name As String = "DrChecks Summary Report"
                         Optional include_timestamp as Boolean = True) As Workbook
    'Creates a new workbook with the provided name and appends with a timestamp as noted.
    Dim combined_workbook As Workbook
    Dim file_name as String
    Set combined_workbook = Workbooks.Add
    Application.DisplayAlerts = False
    With combined_workbook
        .Title = workbook_name
        If include_timestamp Then
            file_name = save_path & "\" & workbook_name & " " & Format(Now(), "YYYY-MM-DD hh-mm-ss") & ".xlsx"
        Else
            file_name = save_path & "\" & workbook_name & ".xlsx"
        End if
        .SaveAs Filename:=file_name, FileFormat:=xlOpenXMLWorkbook
    End With
    Application.DisplayAlerts = True
    Set create_workbook = combined_workbook
End Function

Function verify_projnet_xml(xml_file_path As String) As Boolean
    'Opens an XML to see if the root element is <ProjNet> and returns T/F
    Dim xml_doc As DOMDocument60
    Dim root_element As IXMLDOMElement
    Dim is_valid As Boolean

    Set xml_doc = New DOMDocument60
    xml_doc.validateOnParse = False
    xml_doc.Load xml_file_path
    
    Set root_element = xml_doc.DocumentElement
    If root_element.nodeName = "ProjNet" Then is_valid = True
    verify_projnet_xml = is_valid
End Function

Private Sub perform_on_each_file()

    Dim fd As FileDialog
    Dim fso As New FileSystemObject
    Dim my_folder As String
    Dim a_file As File
    Dim summary_workbook As Workbook
    Dim current_sheet As Worksheet
    
    Application.ScreenUpdating = False
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = "Select Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        my_folder = .SelectedItems(1)
    End With
    
    Set summary_workbook = create_workbook(my_folder)
    i = 1
    On Error Resume Next
    For Each a_file In fso.GetFolder(my_folder).Files
        If fso.GetExtensionName(a_file) = "xml" Then
            If verify_projnet_xml(a_file.Path) = True Then
                'Debug.Print summary_workbook.Sheets.count
                If i = 1 Then
                    Set current_sheet = summary_workbook.Sheets(1)
                Else
                    summary_workbook.Sheets.Add After:=summary_workbook.Sheets(summary_workbook.Sheets.count)
                    Set current_sheet = summary_workbook.Sheets(summary_workbook.Sheets.count)
                End If
                
                import_to_sheet a_file.Path, current_sheet
                i = i + 1
            End If
        End If
        
    Next
    
    summary_workbook.Close SaveChanges:=True
    
    Application.ScreenUpdating = True
    On Error GoTo 0

End Sub

Sub import_to_sheet(ByVal xml_file_path As String, ByVal a_sheet As Worksheet)
    
    Dim st, et As Double
    st = Timer
    
    'Application.ScreenUpdating = False
    
    'Start fresh
    With a_sheet.Cells
        'To make sure you don't keep stacking grouped regions
        If a_sheet.UsedRange.Rows.OutlineLevel >= 1 Then
            .ClearOutline
        End If
        .Delete
    End With
    ActiveWindow.DisplayGridlines = False
    a_sheet.UsedRange.Columns.ColumnWidth = small_column
    
    'xml_file_path = "C:\Users\benst\Documents\_0 Workspace\97 XML Projects\small.xml"
    
    Dim xml_doc As DOMDocument60
    Dim root_element As IXMLDOMElement
    Dim header_info As IXMLDOMNode
    Dim i, header_info_count As Long
    Dim arr As Variant
    
    Set xml_doc = New DOMDocument60
    xml_doc.validateOnParse = False
    xml_doc.Load xml_file_path
    Set root_element = xml_doc.DocumentElement
    
    rename_sheet root_element, a_sheet
    
    Set target_cell = a_sheet.Range(target_cell_address)
    Set user_start_cell = create_project_info_region(root_element, target_cell)
    
    header_row = user_start_cell.Row
    Set comment_header_start_cell = create_user_region_header(user_start_cell)
    Set response_header_start_cell = create_comment_region_header(comment_header_start_cell)
    
    create_response_region_header response_header_start_cell, root_element
    
    Set comment_target_cell = comment_header_start_cell.Offset(1, 0)
    Set data_range = get_all_comment_data(comment_target_cell, root_element)
    
    apply_color_code_conditional_format data_range
    apply_row_lines data_range
    apply_user_region_formats data_range
    
    Set response_data_region = Range(response_header_start_cell.Offset(1, 0), _
                                    Cells(a_sheet.UsedRange.Rows.count, _
                                            a_sheet.UsedRange.Columns.count))
                                            
    apply_response_formatting response_data_region
    
    'Additional Formatting after the header is created
    set_max_RowHeight data_range
    With a_sheet.Rows(header_row)
        .AutoFilter
        .Font.Bold = True
        .VerticalAlignment = xlVAlignBottom
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinous
            .Weight = xlMedium
        End With
    End With

    a_sheet.UsedRange.HorizontalAlignment = xlHAlignLeft
    data_range.WrapText = True
    data_range.VerticalAlignment = xlVAlignTop
    a_sheet.Outline.ShowLevels ColumnLevels:=1
    
    'Application.ScreenUpdating = True
    
    et = Timer
    Debug.Print (et - st) * 1000 & " ms"
    
End Sub

'#########################################################################
'
'                             Layout Methods
'
'#########################################################################

Function create_project_info_region(root_element As IXMLDOMElement, 
                              ByVal target_cell As Range,
                              ByVal a_worksheet as Worksheet = ActiveSheet) as Range
    'Gets the project info from the <DrChecks> element, puts that data into the the 
    'specified target_cell. This function return a Range that indicates the start location
    'of the next region (User Header Region).
    Dim project_info_node as IXMLDOMElement
    Dim project_info_count as Long
    Dim project_info as Variant
    If TypeName(target_cell) = "String" Then
        Set target_cell = a_worksheet.Range(target_cell)
    End If
    Set project_info_node = root_element.selectSingleNode("DrChecks")
    project_info_count = project_info_node.ChildNodes.Length
    ReDim project_info(0 To project_info_count - 1, 0 To 1)
    For i = 0 To project_info_count - 1
        project_info(i, 0) = project_info_node.ChildNodes.Item(i).nodeName
        project_info(i, 1) = CStr(project_info_node.ChildNodes.Item(i).Text)
    Next
    Range(target_cell, target_cell.Offset(header_info_count - 1, 1)) = project_info
    'Formatting of the first column, which has the names of the project info properties:
    With Range(target_cell, target_cell.Offset(header_info_count - 1, 0))
        .Font.Bold = True
    End With
    'Return the cell for the start of the User region header
    Set create_project_info_region = a_worksheet.Cells(target_cell.Offset(header_info_count + 1, 0).Row, 1)
End Function

Function create_user_region_header(ByVal start_cell As Range) As Range
    'Creates a header for the "User Region" where the user can add comments, etc.
    'Returns the cell to start the next region: the Comments Region header
    Dim header_fields, column_widths As Variant
    header_fields = Array("User Notes", "Action Items", "Assignee")
    column_widths = Array(large_column, large_column, medium_column)
    'Plant this information into the Worksheet
    Range(start_cell.Offset(0, LBound(header_fields)), _
          start_cell.Offset(0, UBound(header_fields))) = header_fields
    'Formatting
    For i = 0 To UBound(column_widths)
        start_cell.Offset(0, i).ColumnWidth = column_widths(i)
    Next
    Range(start_cell, start_cell.Offset(0, UBound(arr))).Columns.Group
    'Return the cell address to start the comment region header
    Set create_user_region_header = start_cell.Offset(0, UBound(arr) + 1)
End Function

Function create_comment_region_header(ByVal start_cell As Range) As Range
    'Creates the header for the Comment region. Returns the cell for the start of the
    'evaluations region.
    Dim header_fields, column_widths As Variant
    Dim i as Long
    header_fields = Array("ID", "Comment Status", "Discipline", "Author", "Date", "Comment", "Att.", "Days Open")
    column_widths = Array(small_column, small_column, medium_column, medium_column, _
                            small_column, xlarge_column, xsmall_column, small_column)
    'Place the header data
    Range(start_cell.Offset(0, LBound(header_fields)), start_cell.Offset(0, UBound(header_fields))) = header_fields 
    'Formatting
    For i = 0 To UBound(column_widths)
        start_cell.Offset(0, i).ColumnWidth = column_widths(i)
    Next
    'Return the cell address to start the Response region header (first position is Evaulation)
    Set create_comment_region_header = start_cell.Offset(0, UBound(arr) + 1)
End Function

Function create_response_region_header(ByVal start_cell As Range, root_element As IXMLDOMElement) as Range
    'Creates the header for the response region. This is subdivided into the Evaluations region
    'followed by the Backchecks Region. The create_comments_region_header function returns the start_cell
    'for this function --- which is the same as the start of the Evaluation region. This function
    'returns the start cell of the Backchecks region, which can be used later on.  
    Dim comment_node As IXMLDOMElement
    Dim template_array, column_widths, response_header As Variant
    Dim max_evaluations, max_backchecks, total_response_count As Long
    Dim template_count, total_size, backchecks_start_index as Long
    Dim i, j, k As Long

    'This function basically sets up a template that is repeated as many times as there are
    'evaulations or backchecks found in the XML file.
    template_array = Array("Status", "Author", "Date", "Text", "Att.")
    column_widths = Array(small_column, medium_column, small_column, xlarge_column, xsmall_column)
    template_count = UBound(template_array) + 1

    max_evaluations = get_max_element_count("Comments/comment/evaluations", root_element)
    max_backchecks = get_max_element_count("Comments/comment/backchecks", root_element)
    total_response_count = max_evaluations + max_backchecks
    total_size = total_response_count * (template_count) - 1

    ReDim response_header(0 To total_size)
    For i = 1 To max_evaluations
        For j = 1 To template_count
            response_header(k) = "Eval " & i & " " & template_array(j - 1)
            'start_cell.Offset(0, k) = response_header(k)
            k = k + 1
        Next j
    Next i

    backchecks_start_index = k
    For i = 1 To max_backchecks
        For j = 1 To template_count
            response_header(k) = "BCheck " & i & " " & template_array(j - 1)
            'start_cell.Offset(0, k) = response_header(k)
            k = k + 1
        Next j
    Next i
    
    'Plant all the information in the Worksheet
    For i = LBound(response_header) To UBound(response_header)
        start_cell.Offset(0, i) = response_header(i)
    Next
    
    'We need to do this following formatting here, in order to do the appropriate width sizing.
    With start_cell.EntireRow
        .Font.Bold = True
        .WrapText = True
    End With
    For i = 0 To total_size
        start_cell.Offset(0, i).ColumnWidth = column_widths(i Mod template_count)
    Next
    
    'Group columns. At the very end, we will collapse all outline levels.
    For i = 0 To total_response_count - 1
        Range(start_cell.Offset(0, i * template_count + 1), start_cell.Offset(0, i * template_count + (template_count - 1))).Columns.Group
    Next
    
    create_response_region_header = start_cell.Offset(0, backchecks_start_index)
End Function

'#########################################################################
'
'                  Methods to Build and Inject Comments
'
'#########################################################################

Function get_comment_data(comment_node As IXMLDOMElement)

    Dim arr(6) As Variant
    Dim attach_symbol As String: attach_symbol = ""
   
    arr(0) = comment_node.selectSingleNode("id").Text
    arr(1) = comment_node.selectSingleNode("status").Text
    arr(2) = comment_node.selectSingleNode("Discipline").Text
    arr(3) = comment_node.selectSingleNode("createdBy").Text
    arr(4) = Format(CDate(comment_node.selectSingleNode("createdOn").Text), "YYYY/MM/DD")
    arr(5) = Replace(comment_node.selectSingleNode("commentText").Text, "<br />", vbCrLf)
    If Len(comment_node.selectSingleNode("attachment").Text) > 0 Then has_attach = ChrW(9650)
    arr(6) = has_attach
    
    get_comment_data = arr
    
End Function

Function get_single_response_data(a_node As IXMLDOMElement, Optional is_evaluation As Boolean = True)
    Dim arr(4) As Variant
    Dim attach_symbol As String: attach_symbol = "-"
   
    arr(0) = a_node.selectSingleNode("status").Text
    arr(1) = a_node.selectSingleNode("createdBy").Text
    arr(2) = Format(CDate(a_node.selectSingleNode("createdOn").Text), "YYYY/MM/DD")
    If is_evaluation Then
        arr(3) = a_node.selectSingleNode("evaluationText").Text
    Else
        arr(3) = a_node.selectSingleNode("backcheckText").Text
    End If
    arr(3) = Replace(arr(3), "<br />", vbCrLf)
    If Len(a_node.selectSingleNode("attachment").Text) > 0 Then has_attach = ChrW(9650)
    arr(4) = has_attach
    get_single_response_data = arr
    
End Function

Function get_comment_response_data(comment_node As IXMLDOMElement, _
                                    max_evaluations As Long, _
                                    max_backchecks As Long)
    
    Dim field_count As Integer
    Dim arr As Variant
    Dim item_arr As Variant
    Dim i, j, k As Long

    Dim comment_evaluations As IXMLDOMSelection
    Dim comment_backchecks As IXMLDOMSelection
    Dim response_node As IXMLDOMElement
    
    ' Refer to create_response_region_header() --> template_count
    template_count = 5
    
    Set comment_evaluations = comment_node.selectNodes("evaluations/*")
    Set comment_backchecks = comment_node.selectNodes("backchecks/*")
    
    total_response_count = max_evaluations + max_backchecks
    total_size = total_response_count * (template_count) - 1
    
    ReDim arr(total_size)
    
    For i = 0 To comment_evaluations.Length - 1
        Set response_node = comment_evaluations.Item(i)
        item_arr = get_single_response_data(response_node, is_evaluation:=True)
        For j = 0 To UBound(item_arr)
            arr(k) = item_arr(j)
            k = k + 1
        Next j
    Next i
    k = max_evaluations * template_count
    For i = 0 To comment_backchecks.Length - 1
        Set response_node = comment_backchecks.Item(i)
        item_arr = get_single_response_data(response_node, is_evaluation:=False)
        For j = 0 To UBound(item_arr)
            arr(k) = item_arr(j)
            k = k + 1
        Next j
    Next i
    
    get_comment_response_data = arr
    
End Function

Function get_all_comment_data(ByVal start_cell As Range, root_element As IXMLDOMElement) As Range
    'TODO: pass root_element as an argument
    
    'Set start_cell = Cells(ActiveSheet.UsedRange.Rows.count + 1, 4)
    Dim comment_node As IXMLDOMElement
    Dim max_evaluations As Long
    Dim max_backchecks As Long
    
    Set comments_selection = root_element.selectNodes("Comments/comment")
    total_comments = count_all_comments(root_element)
    max_evaluations = get_max_element_count("Comments/comment/evaluations", root_element)
    max_backchecks = get_max_element_count("Comments/comment/backchecks", root_element)
    
    For i = 0 To total_comments - 1
        ' Insert all comments in Comments region
        Set comment_node = comments_selection.Item(i)
        arr = get_comment_data(comment_node)
        Range(start_cell.Offset(i, 0), start_cell.Offset(i, UBound(arr))) = arr
        start_cell.Offset(i, UBound(arr) + 1) = count_days_open(comment_node)
        
        ' Insert all response (evaluations and backchecks) in Comments region
        resp_arr = get_comment_response_data(comment_node, max_evaluations, max_backchecks)
        Range(start_cell.Offset(i, UBound(arr) + 2), start_cell.Offset(i, UBound(arr) + UBound(resp_arr) + 2)) = resp_arr
    Next
        
    Set get_all_comment_data = Range(start_cell, Cells(ActiveSheet.UsedRange.Rows.count, ActiveSheet.UsedRange.Columns.count))
    
End Function

'#########################################################################
'
'                         Color Coding Methods
'
'#########################################################################

Public Sub apply_color_code_conditional_format(ByVal data_range As Range)

    data_range.FormatConditions.Delete
    first_cell = Replace(data_range(1).Address, "$", "")
    
    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER($E8)=""closed"""
    data_range.FormatConditions(1).Font.Color = color_silver
    data_range.FormatConditions(1).Interior.Color = color_whitesmoke
    
    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""check and resolve"""
    With data_range.FormatConditions(2)
        With .Borders(xlEdgeLeft)
            .Color = 255
        End With
        .Font.Color = 255
        .Interior.Color = 15461375
    End With
    
    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""non-concur"""
    With data_range.FormatConditions(3)
        With .Borders(xlEdgeLeft)
            .Color = 36799
        End With
        .Font.Color = 36799
        .Interior.Color = 13434879
    End With
    
    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""for information only"""
    With data_range.FormatConditions(4)
        With .Borders(xlEdgeLeft)
            .Color = 3506772
        End With
        .Font.Color = 3506772
        .Interior.Color = 14348258
    End With

    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""concur"""
    With data_range.FormatConditions(5)
        With .Borders(xlEdgeLeft)
            .Color = 12611584
        End With
        .Font.Color = 12611584
        .Interior.Color = 16247773
    End With

End Sub

Sub apply_row_lines(ByVal data_range As Range)
    
    Dim paint_range As Range
    Set paint_range = ActiveSheet.Range( _
                Cells(data_range(0).Row, 1), _
                Cells(ActiveSheet.UsedRange.Rows.count, ActiveSheet.UsedRange.Columns.count))
    paint_range.Borders(xlInsideHorizontal).LineStyle = xlNone
    With paint_range.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = color_gainsboro
    End With
    
End Sub

Sub apply_response_formatting(ByVal response_range As Range)
    
    For i = 0 To response_range.Columns.count - 1
        If i Mod 5 = 0 Then
            Set current_column = response_range.Columns(i + 1)
            With current_column.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = color_gainsboro
                .Weight = xlThin
            End With
        
            For Each a_cell In current_column.Cells
                If a_cell.Value = "" Then
                    apply_diagonalx_lines a_cell
                End If
            Next
        End If
    Next

End Sub

Sub apply_diagonalx_lines(ByVal a_cell As Range)
    For i = xlDiagonalDown To xlDiagonalUp
        With a_cell.Borders(i)
            .LineStyle = xlContinuous
            .Color = color_gainsboro
            .Weight = xlThin
        End With
    Next
End Sub

Sub apply_user_region_formats(ByVal data_range As Range)

    Dim user_range As Range
    Dim user_range_header As Range
    Set user_range = ActiveSheet.Range( _
                Cells(data_range(0).Row, 1), _
                Cells(ActiveSheet.UsedRange.Rows.count, 3))
                
    Set user_range_header = Range(Cells(data_range(0).Row - 1, 1), _
                                    Cells(data_range(0).Row - 1, 3))
           
    With user_range_header
        .VerticalAlignment = xlVAlignBottom
        .Interior.Color = color_gold
        .Font.Color = select_contrast_font(.Interior.Color)
    End With
                
    With user_range
        .VerticalAlignment = xlVAlignTop
        .Interior.Color = color_lemonchiffon
        .Font.Color = select_contrast_font(.Interior.Color)
    End With
    
End Sub



