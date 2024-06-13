VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Collection of comments for a given Comments node

' NOTES:
' Autofilter on header is currently turned off in ApplyFormats()
' Added "Doc", "Sheet", "Spec", "Section", between Date and Comment in Class_Initialize()
' Added the four fields above to CommentsToArray()

Const xxs_width As Long = 5
Const xs_width As Long = 8
Const sm_width As Long = 10
Const md_width As Long = 20
Const lg_width As Long = 30
Const xl_width As Long = 50

Const max_row_height As Long = 75

Dim pComments As Collection

Dim pCommentHeaderFields As Variant
Dim pEvaluationHeaderFields As Variant
Dim pBackcheckHeaderFields As Variant
Dim pCommentWidths As Variant
Dim pResponseHeaderFields As Variant
Dim pResponseHeaderWidths As Variant

Dim pHasEvaluations As Boolean
Dim pHasBackchecks As Boolean

Dim pInfoBody As Range
Dim pInfoHeader As Range
Dim pCommentsHeader As Range
Dim pCommentsBody As Range
Dim pEvaluationsHeader As Range
Dim pEvaluationsBody As Range
Dim pBackchecksHeader As Range
Dim pBackchecksBody As Range
Dim pAllResponseHeader As Range
Dim pAllResponseBody As Range
Dim pFullHeader As Range
Dim pFullBody As Range

Property Get InfoBody() As Range
    Set InfoBody = pInfoBody
End Property

Property Get InfoHeader() As Range
    Set InfoHeader = pInfoHeader
End Property

Property Get CommentsHeader() As Range
    Set CommentsHeader = pCommentsHeader
End Property

Property Get CommentsBody() As Range
    Set CommentsBody = pCommentsBody
End Property

Property Get EvaluationsHeader() As Range
    Set EvaluationsHeader = pEvaluationsHeader
End Property

Property Get EvaluationsBody() As Range
    Set EvaluationsBody = pEvaluationsBody
End Property

Property Get BackchecksHeader() As Range
    Set BackchecksHeader = pBackchecksHeader
End Property

Property Get BackchecksBody() As Range
    Set BackchecksBody = pBackchecksBody
End Property

Property Get AllResponseHeader() As Range
    Set AllResponseHeader = pAllResponseHeader
End Property

Property Get AllResponseBody() As Range
    Set AllResponseBody = pAllResponseBody
End Property

Property Get List() As Collection
    Set List = pComments
End Property

Property Get EvaluationHeaders() As Variant
    EvaluationHeaders = pEvaluationHeaderFields
End Property

Property Get BackcheckHeaders() As Variant
    BackcheckHeaders = pBackcheckHeaderFields
End Property

Property Get HasEvaluations() As Boolean
    HasEvaluations = pHasEvaluations
End Property

Property Get HasBackchecks() As Boolean
    HasBackchecks = pHasBackchecks
End Property

Private Sub Class_Initialize()
    ' Initialize empty collection
    Set pComments = New Collection
    
    ' This may need to be changed, based on the needs of the project. These headers are the defaults.
    pCommentHeaderFields = Array("ID", "Status", "Discipline", "Author", "Date", "Source", "Reference", "Sheet", "Spec", "Section", "Comment", "Class", "Att.", "Days Open", "Highest Resp.")
    pCommentWidths = Array(sm_width, sm_width, md_width, md_width, sm_width, md_width, md_width, sm_width, md_width, md_width, xl_width, sm_width, xxs_width, xs_width, sm_width)
    
    ' This is a template array - each individual response should prefix these headers.
    pResponseHeaderFields = Array("Status", "Author", "Date", "Text", "Att.")
    pResponseHeaderWidths = Array(sm_width, md_width, sm_width, xl_width, xs_width)
    
End Sub

Sub CreateFromRootElement(root_node As IXMLDOMElement)
    Dim a_comment As Comment
    Dim comment_node As IXMLDOMElement
    Dim i As Long, j As Long, k As Long
    Dim arr As Variant
    If root_node.SelectNodes("Comments").Length > 0 Then
        For Each comment_node In root_node.SelectNodes("Comments/*")
            Set a_comment = New Comment
            a_comment.CreateFromNode comment_node, i
            pComments.Add a_comment
            i = i + 1
        Next
    End If
    
    ' Load comment descendant evaluations if present.
    If Me.MaxEvaluations > 0 Then
        pHasEvaluations = True
        pEvaluationHeaderFields = CreateResponseHeaders(pResponseHeaderFields, "Eval ", Me.MaxEvaluations)
    End If
     
    ' Load comment descendant backchecks if present.
    If Me.MaxBackchecks > 0 Then
        pHasBackchecks = True
        pBackcheckHeaderFields = CreateResponseHeaders(pResponseHeaderFields, "BCheck ", Me.MaxBackchecks)
    End If
End Sub
 
Function CreateResponseHeaders(response_header_fields As Variant, prefix_string As String, max_responses As Long) As Variant
    Dim total_field_count As Long
    Dim header_arr As Variant
    Dim ith_response As Long, jth_field As Long, arr_index As Long
    
    ' response_header_fields is a Base-0 array of the generic header field names. Multiply by the max_responses to get
    ' the total number of unique fields, however, as we'll use this to dimension the header_arr, that number will need
    ' to subtract one, in order to convert to Base-0
    total_field_count = (UBound(response_header_fields) + 1) * max_responses
    ReDim header_arr(total_field_count - 1)
    
    ' Note that each array here is Base-0
    For ith_response = 0 To max_responses - 1
        For jth_field = 0 To UBound(response_header_fields)
            header_arr(arr_index) = prefix_string & (ith_response + 1) & " " & response_header_fields(jth_field)
            arr_index = arr_index + 1
        Next jth_field
    Next ith_response
    
    'Return the array
    CreateResponseHeaders = header_arr
End Function

Function Item(index As Long)
    Set Item = pComments.Item(index)
End Function

Function HasItems() As Boolean
    If pComments.Count <> 0 Then HasItems = True
End Function

Function Count() As Long
    Count = pComments.Count
End Function

Function MaxEvaluations() As Long
    Dim eval_count As Long
    Dim cmt As Comment
    For Each cmt In pComments
        If cmt.EvaluationsCount > eval_count Then eval_count = cmt.EvaluationsCount
    Next
    MaxEvaluations = eval_count
End Function

Function MaxBackchecks() As Long
    Dim eval_count As Long
    Dim cmt As Comment
    For Each cmt In pComments
        If cmt.BackchecksCount > eval_count Then eval_count = cmt.BackchecksCount
    Next
    MaxBackchecks = eval_count
End Function

Function OffsetToEvaluations() As Long
    ' Offset relative to the start_cell for the comment region
    ' Add 1 to adjust from Base-0 to Base-1 value
    
    OffsetToEvaluations = UBound(pCommentHeaderFields) + 1
End Function

Function OffsetToBackchecks() As Long
    ' Offset relative to the start_cell for the comment region
    ' Add 1 twise to adjust from Base-0 to Base-1 value; UBound is zero-based.
    
    OffsetToBackchecks = UBound(pCommentHeaderFields) + UBound(pEvaluationHeaderFields) + 2
End Function

Function HeaderCount() As Long
    HeaderCount = UBound(pCommentHeaderFields) + 1
End Function

Function CommentHeaderRange(start_cell As Range) As Range
    ' Offset of 0,0 is no offset, so offset is in a sense Base-0, which is why we subtract 1 from Me.HeaderCount
    Set pCommentsHeader = Range(start_cell, start_cell.Offset(0, Me.HeaderCount - 1))
    Set CommentHeaderRange = pCommentsHeader
End Function

Function CommentRange(start_cell As Range) As Range
    ' Offset of 0,0 is no offset, so offset is in a sense Base-0, which is why we subtract 1 from Me.HeaderCount
    Set pCommentsBody = Range(start_cell.Offset(1, 0), start_cell.Offset(Me.Count, Me.HeaderCount - 1))
    Set CommentRange = pCommentsBody
End Function

Function CommentsToArray() As Variant
    Dim aComment As Comment
    Dim arr As Variant
    Dim i As Long
    ReDim arr(Me.Count - 1, Me.HeaderCount - 1)
    i = 0
    For Each aComment In pComments
        arr(i, 0) = aComment.ID
        arr(i, 1) = aComment.status
        arr(i, 2) = aComment.Discipline
        arr(i, 3) = aComment.CreatedBy
        arr(i, 4) = Format(aComment.CreatedOn, "mm/dd/yy")
        
        '"Source", "Ref", "Sheet", "Spec", "Section",
        arr(i, 5) = aComment.DocType & " " & aComment.DocRef
        arr(i, 6) = Find_StandardReferences(aComment.DocType, aComment.DocRef, aComment.Text)
        arr(i, 7) = Find_SheetNumbers(aComment.SheetRef, aComment.Detail, aComment.DocRef, aComment.Text)
        arr(i, 8) = Find_SpecSections(aComment.DocRef, aComment.Spec, aComment.Text)
        arr(i, 9) = Find_ReportSectionNumbers(aComment.DocType, aComment.DocRef, aComment.Text)
        
        arr(i, 10) = aComment.Text
        
        arr(i, 12) = aComment.Attachment
        arr(i, 13) = aComment.DaysOpen
        arr(i, 14) = aComment.HighestStatus
        i = i + 1
    Next
    CommentsToArray = arr
End Function

Function ResponseHeaderFieldCount() As Long
    ResponseHeaderFieldCount = UBound(pResponseHeaderFields) + 1
End Function

Function EvaluationsHeaderRange(start_cell As Range) As Range
    Set pEvaluationsHeader = Range(start_cell, start_cell.Offset(0, (Me.ResponseHeaderFieldCount) * (Me.MaxEvaluations) - 1))
    Set EvaluationsHeaderRange = pEvaluationsHeader
End Function

Function EvaluationsRange(start_cell As Range) As Range
    Set pEvaluationsBody = Range(start_cell.Offset(1, 0), start_cell.Offset(Me.Count, (Me.ResponseHeaderFieldCount) * (Me.MaxEvaluations) - 1))
    Set EvaluationsRange = pEvaluationsBody
End Function

Function EvaluationsToArray() As Variant
    Dim eval_arr As Variant
    Dim a_comment As New Comment
    Dim an_evaluation As New Evaluation
    Dim ith_comment As Long, jth_field As Long
    Dim header_field_count As Long
    
    header_field_count = UBound(pResponseHeaderFields) + 1
    ReDim eval_arr(Me.Count, UBound(pEvaluationHeaderFields))
    For Each a_comment In pComments
        For jth_field = 0 To a_comment.EvaluationsCount - 1
            Set an_evaluation = a_comment.EvaluationsList.Item(jth_field + 1)
            eval_arr(ith_comment, jth_field * header_field_count + 0) = StrConv(an_evaluation.status, vbProperCase)
            eval_arr(ith_comment, jth_field * header_field_count + 1) = an_evaluation.CreatedBy
            eval_arr(ith_comment, jth_field * header_field_count + 2) = Format(an_evaluation.CreatedOn, "mm/dd/yyyy")
            eval_arr(ith_comment, jth_field * header_field_count + 3) = an_evaluation.Text
            eval_arr(ith_comment, jth_field * header_field_count + 4) = an_evaluation.Attachment
        Next
        ith_comment = ith_comment + 1
    Next
    EvaluationsToArray = eval_arr
End Function

Function BackchecksHeaderRange(start_cell As Range) As Range
    Set pBackchecksHeader = Range(start_cell, start_cell.Offset(0, (ResponseHeaderFieldCount) * (Me.MaxBackchecks) - 1))
    Set BackchecksHeaderRange = pBackchecksHeader
End Function

Function BackchecksRange(start_cell As Range) As Range
    Set pBackchecksBody = Range(start_cell.Offset(1, 0), start_cell.Offset(Me.Count, (ResponseHeaderFieldCount) * (Me.MaxBackchecks) - 1))
    Set BackchecksRange = pBackchecksBody
End Function

Function BackchecksToArray() As Variant
    Dim backcheck_arr As Variant
    Dim a_comment As New Comment
    Dim a_backcheck As New Backcheck
    Dim ith_comment As Long, jth_field As Long
    Dim header_field_count As Long
    
    header_field_count = UBound(pResponseHeaderFields) + 1
    ReDim backcheck_arr(Me.Count, UBound(pBackcheckHeaderFields))
    For Each a_comment In pComments
        For jth_field = 0 To a_comment.BackchecksCount - 1
            Set a_backcheck = a_comment.BackchecksList.Item(jth_field + 1)
            backcheck_arr(ith_comment, jth_field * header_field_count + 0) = StrConv(a_backcheck.status, vbProperCase)
            backcheck_arr(ith_comment, jth_field * header_field_count + 1) = a_backcheck.CreatedBy
            backcheck_arr(ith_comment, jth_field * header_field_count + 2) = Format(a_backcheck.CreatedOn, "mm/dd/yyyy")
            backcheck_arr(ith_comment, jth_field * header_field_count + 3) = a_backcheck.Text
            backcheck_arr(ith_comment, jth_field * header_field_count + 4) = a_backcheck.Attachment
        Next
        ith_comment = ith_comment + 1
    Next
    BackchecksToArray = backcheck_arr
End Function

Function PasteData(target_cell As Range) As Variant
    'Paste all the proeprties into an Excel Worksheet, return range of offset specified from end
    Dim header_range As Range
    Dim comment_range As Range
    
    Dim eval_header_range As Range
    Dim eval_range As Range
    
    Dim bcs_header_range As Range
    Dim bcs_range As Range

    Dim i As Long
    
    Application.ScreenUpdating = False
    
    ' Set ranges for data, then paste in the data
    Set header_range = CommentHeaderRange(target_cell)
    Set comment_range = CommentRange(target_cell)
    header_range = pCommentHeaderFields
    comment_range = Me.CommentsToArray()
    
    Set pFullHeader = pCommentsHeader
    Set pFullBody = pCommentsBody
    
    If Me.HasEvaluations Then
        Set eval_header_range = EvaluationsHeaderRange(target_cell.Offset(0, Me.OffsetToEvaluations))
        Set eval_range = EvaluationsRange(target_cell.Offset(0, Me.OffsetToEvaluations))
        eval_header_range = Me.EvaluationHeaders
        eval_range = Me.EvaluationsToArray()
        
        Set pAllResponseHeader = eval_header_range
        Set pAllResponseBody = eval_range
        
        Set pFullHeader = Union(pFullHeader, eval_header_range)
        Set pFullBody = Union(pFullBody, eval_range)
    End If
    
    If Me.HasBackchecks Then
        Set bcs_header_range = BackchecksHeaderRange(target_cell.Offset(0, Me.OffsetToBackchecks))
        Set bcs_range = BackchecksRange(target_cell.Offset(0, Me.OffsetToBackchecks))
        bcs_header_range = Me.BackcheckHeaders
        bcs_range = Me.BackchecksToArray()

        Set pAllResponseHeader = Union(eval_header_range, bcs_header_range)
        Set pAllResponseBody = Union(eval_range, bcs_range)
        
        Set pFullHeader = Union(pFullHeader, bcs_header_range)
        Set pFullBody = Union(pFullBody, bcs_range)
    End If
    
    Set PasteData = header_range
    Application.ScreenUpdating = True
End Function

Sub ApplyXLines(ByVal a_cell As Range)
    For i = xlDiagonalDown To xlDiagonalUp
        With a_cell.Borders(i)
            .LineStyle = xlContinuous
            .Color = GAINSBORO
            .Weight = xlThin
        End With
    Next
End Sub

Public Sub ApplyConditionalFormats(ByVal data_range As Range)

    data_range.FormatConditions.Delete
    first_cell = Replace(data_range(1).Address, "$", "")
    first_status_cell = Replace(data_range(2).Address, "$", "")
    
    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER($" & first_status_cell & ")=""closed"""
    data_range.FormatConditions(1).Font.Color = SILVER
    data_range.FormatConditions(1).Interior.Color = WHITESMOKE
    
    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""check and resolve"""
    With data_range.FormatConditions(2)
        With .Borders(xlEdgeLeft)
            .Color = TOMATO
        End With
        .Font.Color = TOMATO
        .Interior.Color = MISTYROSE
    End With
    
    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""non-concur"""
    With data_range.FormatConditions(3)
        With .Borders(xlEdgeLeft)
            .Color = DARKGOLDENROD
        End With
        .Font.Color = DARKGOLDENROD
        .Interior.Color = LEMONCHIFFON
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

Public Sub FormatHighest(ByVal data_range As Range)

    Dim cnt As Long
    cnt = data_range.FormatConditions.Count
    For i = cnt To 2 Step -1
        data_range.FormatConditions(i).Delete
    Next

    first_cell = Replace(data_range(1).Address, "$", "")
    first_status_cell = Replace(data_range(2).Address, "$", "")

    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""check and resolve"""
    With data_range.FormatConditions(2)
        .Interior.Color = webcolors.ORANGERED
        .Font.Color = webcolors.ContrastText(.Interior.Color)
    End With

    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""non-concur"""
    With data_range.FormatConditions(3)
        .Interior.Color = GOLD
        .Font.Color = webcolors.ContrastText(.Interior.Color)
    End With

    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""for information only"""
    With data_range.FormatConditions(4)
        '.Interior.Color = 3506772
        .Interior.Color = webcolors.MEDIUMSEAGREEN
        .Font.Color = webcolors.ContrastText(.Interior.Color)
    End With

    data_range.FormatConditions.Add Type:=xlExpression, Formula1:="=LOWER(" & first_cell & ")=""concur"""
    With data_range.FormatConditions(5)
        '.Interior.Color = 12611584
        .Interior.Color = webcolors.DODGERBLUE
        .Font.Color = webcolors.ContrastText(.Interior.Color)
    End With

End Sub




Function ApplyFormats()
    ' Reboot Sheet Styles
    Application.ScreenUpdating = False
    With ActiveSheet
        ' To make sure you don't keep stacking grouped regions
        If .UsedRange.Rows.OutlineLevel >= 1 Then
            .Cells.ClearOutline
        End If
        Cells.Columns.ColumnWidth = xs_width
    End With
    ActiveWindow.DisplayGridlines = False
    
    ' Start to format
    With pCommentsHeader.EntireRow
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignBottom
    End With
    
    With pFullHeader
        .Borders(xlEdgeBottom).LineStyle = xlNone
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = 12611584
        End With
    End With
    
    For i = 0 To UBound(pCommentWidths)
        CommentsHeader(i + 1).ColumnWidth = pCommentWidths(i)
    Next

    With pCommentsBody
        ' The Comment Text column
        With .Columns(11)
            .WrapText = True
            .Font.Size = 8
        End With
        For Each a_row In .Rows
            If a_row.RowHeight > max_row_height Then a_row.RowHeight = max_row_height
        Next
    End With
    
    With pCommentsBody.Columns(15)
        .WrapText = True
        For Each a_cell In .Rows
            With a_cell.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = GAINSBORO
                .Weight = xlThin
            End With
            If a_cell = Empty Then ApplyXLines a_cell
        Next
    End With
    Set pHighestResponseBody = pCommentsBody.Columns(15)
    Debug.Print pHighestResponseBody.Address
    
    With pFullBody
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = GAINSBORO
        End With
    End With

    Me.ApplyConditionalFormats pFullBody
    
    
    
    
    If Me.HasEvaluations Then
        For i = 0 To pAllResponseHeader.Columns.Count - 1
            pAllResponseHeader(i + 1).ColumnWidth = pResponseHeaderWidths(i Mod (UBound(pResponseHeaderFields) + 1))
            ' Find the column that has the comments --- in this case it is 4 (or 3 in Base-0)
            If i Mod ((UBound(pResponseHeaderFields) + 1)) = 3 Then
                pAllResponseHeader(i + 1).EntireColumn.WrapText = True
                pAllResponseBody(i + 1).EntireColumn.Font.Size = 8
            End If

            ' Find the first column and then group the next through last columns
            If i Mod ((UBound(pResponseHeaderFields) + 1)) = 0 Then
                pAllResponseBody(i + 1).Offset(0, 0).EntireColumn.WrapText = True
                On Error Resume Next
                For Each a_cell In ActiveSheet.Range(pAllResponseBody(i + 1).Offset(0, 0), pAllResponseBody(i + 1).Offset(Me.Count - 1, 0)).Cells
                    If a_cell = Empty Then ApplyXLines a_cell
                    With a_cell.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Color = GAINSBORO
                        .Weight = xlThin
                    End With
                Next
                
                ' There are five total fields, so we want to group the last four of each response
                ActiveSheet.Range(pAllResponseBody(i + 1).Offset(0, 1), pAllResponseBody(i + 1).Offset(0, 4)).Group
            End If
        Next
        With pAllResponseBody
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignTop
        End With
    End If
    
    Dim refRng As Range
    Set refRng = Union(pCommentsHeader(6), pCommentsHeader(7), pCommentsHeader(8), pCommentsHeader(9), pCommentsHeader(10))
    With refRng.EntireColumn
        .Group
        .Font.Size = 9
        .WrapText = True
    End With
    
    dxreview.GroupedColumnTriangles
    
    ActiveSheet.Range(ALLCOMMENTSTARGETCELL).EntireRow.Font.Size = 9
    
    Me.FormatHighest pCommentsBody.Columns(15)
    
    'pFullHeader.AutoFilter
    
    ' Collapse all "outline" groups
    ActiveSheet.Outline.ShowLevels ColumnLevels:=1
    Application.ScreenUpdating = True
End Function



