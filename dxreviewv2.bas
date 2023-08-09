Attribute VB_Name = "dxreviewv2"

Const target_cell_address = "D1"

Const xxlarge_column = 50
Const xlarge_column = 40
Const large_column = 30
Const medium_column = 20
Const small_column = 10
Const xsmall_column = 5

Const max_row_height = 75

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

Function get_nodetype_from_enum(node_type_number)
   
    Dim arr As Variant
    arr = Array("NODE_ELEMENT", _
                "NODE_ATTRIBUTE", _
                "NODE_TEXT", _
                "NODE_CDATA_SECTION", _
                "NODE_ENTITY_REFERENCE", _
                "NODE_ENTITY", _
                "NODE_PROCESSING_INSTRUCTION", _
                "NODE_COMMENT", _
                "NODE_DOCUMENT", _
                "NODE_DOCUMENT_TYPE", _
                "NODE_DOCUMENT_FRAGMENT", _
                "NODE_NOTATION")
    get_nodetype_from_enum = arr(node_type_number - 1)
    
End Function

Private Sub apply_max_row_height()
    For Each r In ActiveSheet.UsedRange.Rows
        If r.RowHeight > max_row_height Then r.RowHeight = max_row_height
    Next
End Sub

Function count_all_comments(root_element As IXMLDOMElement)
    ' COMPLETED
    Dim comments_selection As IXMLDOMSelection
    Set comments_selection = root_element.selectNodes("Comments/comment")
    count_all_comments = comments_selection.Length
End Function

Function count_all_evaluations(root_element As IXMLDOMElement)
    ' COMPLETED
    Dim evaluations_selection As IXMLDOMSelection
    Dim i As Long
    Dim total_count As Long
    Set evaluations_selection = root_element.selectNodes("Comments/comment/evaluations")
    For i = 0 To evaluations_selection.Length - 1
        total_count = total_count + evaluations_selection(i).ChildNodes.Length
    Next
    count_all_evaluations = total_count
End Function

Function count_all_backchecks(root_element As IXMLDOMElement)
    ' COMPLETED
    Dim backchecks_selection As IXMLDOMSelection
    Dim i As Long
    Dim total_count As Long
    Set backchecks_selection = root_element.selectNodes("Comments/comment/backchecks")
    For i = 0 To backchecks_selection.Length - 1
        total_count = total_count + backchecks_selection(i).ChildNodes.Length
    Next
    count_all_backchecks = total_count
End Function

Function get_max_evaluations(root_element As IXMLDOMElement)
    ' COMPLETED
    Dim evaluations_selection As IXMLDOMSelection
    Dim max_count As Long
    Set evaluations_selection = root_element.selectNodes("Comments/comment/evaluations")
    For i = 0 To evaluations_selection.Length - 1
        If max_count < evaluations_selection(i).ChildNodes.Length Then
            max_count = evaluations_selection(i).ChildNodes.Length
        End If
    Next
    get_max_evaluations = max_count
End Function

Function get_max_backchecks(root_element As IXMLDOMElement)
    ' COMPLETED
    Dim backchecks_selection As IXMLDOMSelection
    Dim max_count As Long
    Set backchecks_selection = root_element.selectNodes("Comments/comment/backchecks")
    For i = 0 To backchecks_selection.Length - 1
        If max_count < backchecks_selection(i).ChildNodes.Length Then
            max_count = backchecks_selection(i).ChildNodes.Length
        End If
    Next
    get_max_backchecks = max_count
End Function

Function count_comment_evaluations(ByVal a_node As IXMLDOMNode) As Long
    ' COMPLETED
    Set evaulations = a_node.selectNodes("evaluations")
    count_comment_evaluations = evaulations.Item(0).ChildNodes.Length
End Function

Function count_comment_backchecks(ByVal a_node As IXMLDOMNode) As Long
    ' COMPLETED
    Set backchecks = a_node.selectNodes("backchecks")
    count_comment_backchecks = backchecks.Item(0).ChildNodes.Length
End Function

Function count_days_open(comment_node As IXMLDOMElement) As Long

    Dim comment_status As String
    Dim comment_created_date As Date

    Dim evaluations_selection As IXMLDOMSelection
    Dim evaluations_count As Long
    Dim last_evaluation As IXMLDOMElement
    Dim last_evaluation_date As Date
    
    Dim backchecks_section As IXMLDOMSelection
    Dim last_backcheck As IXMLDOMElement
    Dim backchecks_count As Long
    Dim last_backcheck_date As Date
    
    comment_status = LCase(comment_node.selectSingleNode("status").Text)
    comment_created_date = CDate(comment_node.selectSingleNode("createdOn").Text)
    If comment_status = "closed" Then
        Set backchecks_section = comment_node.selectNodes("backchecks/*")
        backchecks_count = backchecks_section.Length
        Set last_backcheck = backchecks_section.Item(backchecks_count - 1)
        last_backcheck_date = CDate(last_backcheck.selectSingleNode("createdOn").Text)
        count_days_open = DateDiff("d", comment_created_date, last_backcheck_date)
    Else
        count_days_open = DateDiff("d", comment_created_date, Now())
    End If
    
End Function

Sub rename_sheet(root_element As IXMLDOMElement, a_sheet As Worksheet)
    Dim arr As Variant
    'Create array of characters that are not permitted in worksheet names
    arr = Array("/", "\", "?", "*", ":", "[", "]")
    
    'The worksheet name will be the Dr Checks <ReviewName>
    review_name = root_element.selectSingleNode("DrChecks/ReviewName").Text
    
    'Make sure the name string does not exceed Excel's length limit
    If Len(review_name) > 31 Then review_name = Left(review_name, 30)
    For i = LBound(arr) To UBound(arr)
        review_name = Replace(review_name, arr(i), "")
    Next
    On Error GoTo dump
    a_sheet.Name = review_name
dump:
End Sub

Public Function rgb_to_hsb(rgb_arr As Variant)

    Dim color_scale As Integer: color_scale = 255

    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim c_max As Double
    Dim c_min As Double
    Dim c_delta As Double
    Dim arr(2) As Integer

    Dim hue As Integer
    Dim sat As Integer
    Dim bright As Integer

    r = rgb_arr(0) / color_scale
    g = rgb_arr(1) / color_scale
    b = rgb_arr(2) / color_scale
    
    c_max = WorksheetFunction.Max(r, g, b)
    c_min = WorksheetFunction.Min(r, g, b)
    
    c_delta = c_max - c_min
    
    If c_max = r And g >= b Then
        hue = 60 * (g - b) / c_delta
    ElseIf c_max = r And g < b Then
        hue = 60 * (g - b) / c_delta + 360
    ElseIf c_max = g Then
        hue = 60 * (b - r) / c_delta + 120
    ElseIf c_max = b Then
        hue = 60 * (r - g) / c_delta + 240
    Else
        hue = 0
    End If
    
    If c_max <> 0 Then sat = c_delta / c_max * 100
    
    bright = c_max * 100
    
    arr(0) = Int(hue)
    arr(1) = Int(sat)
    arr(2) = Int(bright)

    rgb_to_hsb = arr

End Function

Public Function hsb_to_rgb(hsb_arr As Variant)
    'Note H is 360 scale, S and V or B on 100 scale
    
    Dim color_scale As Integer: color_scale = 255
    
    Dim chroma As Double
    Dim x As Double
    Dim m As Double
        
    Dim arr As Variant

    hue = hsb_arr(0)
    sat = hsb_arr(1) / 100
    bright = hsb_arr(2) / 100

    chroma = bright * sat
    x = chroma * (1 - Abs(hue / 60 - Int(hue / 60) - 1))
    m = bright - chroma
    
    If hue >= 0 And hue < 60 Then
        arr = Array(chroma, x, 0)
    ElseIf hue >= 60 And hue < 120 Then
        arr = Array(x, chroma, 0)
    ElseIf hue >= 120 And hue < 180 Then
        arr = Array(0, chroma, x)
    ElseIf hue >= 180 And hue < 240 Then
        arr = Array(0, x, chroma)
    ElseIf hue >= 240 And hue < 300 Then
        arr = Array(x, 0, chroma)
    ElseIf hue >= 300 And hue < 360 Then
        arr = Array(chroma, 0, x)
    Else
        arr = Array(0, 0, 0)
    End If
    
    arr(0) = Int((arr(0) + m) * color_scale)
    arr(1) = Int((arr(1) + m) * color_scale)
    arr(2) = Int((arr(2) + m) * color_scale)
    
    hsb_to_rgb = arr
    
End Function

Public Function long_to_rgb(a_long As Long)
    ReDim arr(0 To 2)
    b = a_long \ 65536
    g = (a_long - b * 65536) \ 256
    r = a_long - b * 65536 - g * 256
    arr(0) = r: arr(1) = g: arr(2) = b
    long_to_rgb = arr
End Function

Public Function rgb_to_long(rgb_arr As Variant) As Long
    rgb_to_long = RGB(rgb_arr(0), rgb_arr(1), rgb_arr(2))
End Function

Public Function apply_contrasting_font_color(background_color As Long)
    'Based on W3.org visibility recommendations:
    'https://www.w3.org/TR/AERT/#color-contrast
    
    Dim arr As Variant
    Dim color_constant As Long
    Dim color_brightness As Double
    
    arr = long_to_rgb(background_color)
    color_brightness = (0.299 * arr(0) + 0.587 * arr(1) + 0.114 * arr(2)) / 255
    If color_brightness > 0.55 Then color_constant = vbBlack Else color_constant = vbWhite
    
    apply_contrasting_font_color = color_constant
    
End Function



'#########################################################################
'
'                             MAIN METHOD
'
'#########################################################################

Sub MAIN()

    perform_on_each_file

End Sub

Function create_workbook(save_path, Optional wb_name As String = "DrChecks Summary Report ") As Workbook
    
    Dim combined_workbook As Workbook
    Set combined_workbook = Workbooks.Add
    Application.DisplayAlerts = False
    With combined_workbook
        .Title = "Combined"
        .SaveAs Filename:=save_path & "\" & wb_name & Format(Now(), "YYYY-MM-DD hh-mm-ss") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End With
    Application.DisplayAlerts = True
    Set create_workbook = combined_workbook
    
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
'    If TypeName(xml_file_path) = "Range" Then
'        xml_doc.Load xml_file_path.Value
'    Else
'        xml_doc.Load xml_file_path
'    End If
    xml_doc.Load xml_file_path
    Set root_element = xml_doc.DocumentElement
    
    rename_sheet root_element, a_sheet
    
    Set target_cell = a_sheet.Range(target_cell_address)
    Set user_start_cell = get_project_metadata(root_element, target_cell)
    
    header_row = user_start_cell.Row
    Set comment_header_start_cell = create_header_user_region(user_start_cell)
    Set response_header_start_cell = create_header_comment_region(comment_header_start_cell)
    
    create_header_response_region response_header_start_cell, root_element
    
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
    apply_max_row_height
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


Function get_project_metadata(root_element As IXMLDOMElement, ByVal target_cell As Range)
    
    If TypeName(target_cell) = "String" Then
        Set target_cell = ActiveSheet.Range(target_cell)
    End If
    
    Set header_info = root_element.selectSingleNode("DrChecks")
    
    header_info_count = header_info.ChildNodes.Length
    
    ReDim arr(0 To header_info_count - 1, 0 To 1)
    For i = 0 To header_info_count - 1
        arr(i, 0) = header_info.ChildNodes.Item(i).nodeName
        arr(i, 1) = CStr(header_info.ChildNodes.Item(i).Text)
    Next
    
    Range(target_cell, target_cell.Offset(header_info_count - 1, 1)) = arr
    
    'Formatting
    With Range(target_cell, target_cell.Offset(header_info_count - 1, 0))
        .Font.Bold = True
    End With
    
    'Return the cell for the start of the User region header
    Set get_project_metadata = Cells(target_cell.Offset(header_info_count + 1, 0).Row, 1)
End Function

Function create_header_user_region(ByVal start_cell As Range) As Range
    Dim arr As Variant
    Dim column_widths As Variant
    arr = Array("User Notes", "Action Items", "Assignee")
    column_widths = Array(large_column, large_column, medium_column)
    For i = LBound(arr) To UBound(arr)
        start_cell.Offset(0, i) = arr(i)
    Next
    
    'Formatting
    For i = 0 To UBound(column_widths)
        start_cell.Offset(0, i).ColumnWidth = column_widths(i)
    Next
    Range(start_cell, start_cell.Offset(0, UBound(arr))).Columns.Group
    
    'Return the cell address to start the comment region header
    Set create_header_user_region = start_cell.Offset(0, UBound(arr) + 1)
End Function

Function create_header_comment_region(ByVal start_cell As Range) As Range
    Dim arr As Variant
    Dim column_widths As Variant
    arr = Array("ID", "Comment Status", "Discipline", "Author", "Date", "Comment", "Att.", "Days Open")
    column_widths = Array(small_column, small_column, medium_column, medium_column, _
                            small_column, xlarge_column, xsmall_column, small_column)
                            
    For i = LBound(arr) To UBound(arr)
        start_cell.Offset(0, i) = arr(i)
    Next
    
    'Formatting
    For i = 0 To UBound(column_widths)
        start_cell.Offset(0, i).ColumnWidth = column_widths(i)
    Next

    'Return the cell address to start the response region header
    Set create_header_comment_region = start_cell.Offset(0, UBound(arr) + 1)
End Function

Sub create_header_response_region(ByVal start_cell As Range, root_element As IXMLDOMElement)
    ' WARNING: Mix of 0- and 1-based operations
    
    Dim comment_node As IXMLDOMElement
    
    Dim prototype_array As Variant
    Dim column_widths As Variant
    Dim max_evaluations As Long
    Dim max_backchecks As Long
    Dim total_response_count As Long
    Dim i, j, k As Long
    
    Dim arr As Variant
    
    prototype_array = Array("Status", "Author", "Date", "Text", "Att.")
    column_widths = Array(small_column, medium_column, small_column, xlarge_column, xsmall_column)
    prototype_count = UBound(prototype_array) + 1
    
    max_evaluations = get_max_evaluations(root_element)
    max_backchecks = get_max_backchecks(root_element)
    total_response_count = max_evaluations + max_backchecks
    total_size = total_response_count * (prototype_count) - 1
    backchecks_start_index = max_evaluations * prototype_count

    ReDim arr(0 To total_size)
    
    For i = 1 To max_evaluations
        For j = 1 To prototype_count
            arr(k) = "Eval " & i & " " & prototype_array(j - 1)
            start_cell.Offset(0, k) = arr(k)
            k = k + 1
        Next j
    Next i

    For i = 1 To max_backchecks
        For j = 1 To prototype_count
            arr(k) = "BCheck " & i & " " & prototype_array(j - 1)
            start_cell.Offset(0, k) = arr(k)
            k = k + 1
        Next j
    Next i
    
    For i = LBound(arr) To UBound(arr)
        start_cell.Offset(0, i) = arr(i)
    Next
    
    'Formatting
    With start_cell.EntireRow
        .Font.Bold = True
        .WrapText = True
    End With
    For i = 0 To total_size
        start_cell.Offset(0, i).ColumnWidth = column_widths(i Mod prototype_count)
    Next
    
    'Group columns
    For i = 0 To total_response_count - 1
        Range(start_cell.Offset(0, i * prototype_count + 1), start_cell.Offset(0, i * prototype_count + (prototype_count - 1))).Columns.Group
    Next
    
End Sub


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
    
    ' Refer to create_header_response_region() --> prototype_count
    prototype_count = 5
    
    Set comment_evaluations = comment_node.selectNodes("evaluations/*")
    Set comment_backchecks = comment_node.selectNodes("backchecks/*")
    
    total_response_count = max_evaluations + max_backchecks
    total_size = total_response_count * (prototype_count) - 1
    
    ReDim arr(total_size)
    
    For i = 0 To comment_evaluations.Length - 1
        Set response_node = comment_evaluations.Item(i)
        item_arr = get_single_response_data(response_node, is_evaluation:=True)
        For j = 0 To UBound(item_arr)
            arr(k) = item_arr(j)
            k = k + 1
        Next j
    Next i
    k = max_evaluations * prototype_count
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
    max_evaluations = get_max_evaluations(root_element)
    max_backchecks = get_max_backchecks(root_element)
    
    
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
        .Font.Color = apply_contrasting_font_color(.Interior.Color)
    End With
                
    With user_range
        .VerticalAlignment = xlVAlignTop
        .Interior.Color = color_lemonchiffon
        .Font.Color = apply_contrasting_font_color(.Interior.Color)
    End With
    
End Sub


Function verify_projnet_xml(xml_file_path As String) As Boolean

    Dim xml_doc As DOMDocument60
    Dim root_element As IXMLDOMElement
    Dim is_valid As Boolean: is_valid = False
    
    Set xml_doc = New DOMDocument60
    xml_doc.validateOnParse = False
    xml_doc.Load xml_file_path
    
    Set root_element = xml_doc.DocumentElement
    If root_element.nodeName = "ProjNet" Then is_valid = True
    'If root_element.selectSingleNode("DrChecks").nodeName = "DrChecks" Then is_valid = True

    verify_projnet_xml = is_valid
    
End Function
