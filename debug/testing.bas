Attribute VB_Name = "testing"
'' Module requires references to:
''    Microsoft XML, v6.0 (msxml6.dll) - XML parsing functions
''    Microsoft Scripting Runtime (scrrun.dll) - Dictionaries
'
'Const path_concept = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\concept.xml"
'Const path_no_responses = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\no responses.xml"
'Const path_books = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\books.xml"
'Const path_no_comments = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\no comments.xml"
'
'Public Function VBA_MOD(num, div)
'    VBA_MOD = num Mod div
'End Function
'
'
'Private Sub Test_EmptyRange()
'
'    Dim a_range As Range: Set a_range = Range("A1")
'    If a_range Is Nothing Then Debug.Print True Else Debug.Print False
'End Sub
'Private Sub Test_Reboots()
''    With ActiveSheet
''        Debug.Print .UsedRange.Address
''
''        Debug.Print .UsedRange.Rows.OutlineLevel
''        .UsedRange.ClearOutline
''
''    End With
'    Cells.Delete
'
'End Sub
'Private Sub TestVBA_Mod()
'
'    nums = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
'    arr = Array(0, 1, 2, 3, 4)
'    divis = UBound(arr) + 1
'    For i = 0 To UBound(nums)
'        Debug.Print i, nums(i) Mod divis
'    Next
'
'End Sub
'
'Private Sub Test_RangeCount()
'
'    Set a_range = Range("A1:C2")
'    Debug.Print a_range.Columns.Count
'    Debug.Print a_range.Rows.Count
'    Debug.Print a_range.Cells.Count
'
'End Sub
'
'Private Sub Test_Dictionary()
'    Dim pStatuses As New Dictionary
'    Dim a_val As Variant
'    pStatuses.Add "concur", 0
'    pStatuses.Add "for information only", 1
'    pStatuses.Add "non-concur", 2
'    pStatuses.Add "check and resolve", 3
'    pStatuses.Add 0, "concur"
'    pStatuses.Add 1, "for information only"
'    pStatuses.Add 2, "non-concur"
'    pStatuses.Add 3, "check and resolve"
'
'    a_val = "non-concur"
'
'    Debug.Print pStatuses(a_val), pStatuses(a_val) > 3
'    Debug.Print pStatuses(2)
'    Debug.Print pStatuses("check and resolve")
'End Sub
'
'
'
'Public Function VerifyProjNetRoot(ByVal root As IXMLDOMElement) As Boolean
'    ' Return True if the XML file is a Dr Checks/ProjNet document
'    If root Is Nothing Then
'        VerifyProjNetRoot = False
'    ElseIf root.nodeName = "ProjNet" Then
'        VerifyProjNetRoot = True
'    End If
'End Function
'
'Public Function GatherNodeChildren(ByVal a_node As IXMLDOMElement) As Variant
'
'    If a_node Is Nothing Then
'        GatherNodeChildren = -1
'    Else
'        GatherNodeChildren = a_node.ChildNodes.Length
'    End If
'
'End Function
'
'
'Private Sub DrChecksTest()
'    Dim root As IXMLDOMElement
'    Dim drchecksnode As IXMLDOMElement
'    Dim project_info As New ProjectInfo
'    Dim project_info_offset As Range
'
'    Set root = GetRootFromXML(path_concept)
'    Set drchecksnode = root.SelectSingleNode("DrChecks")
'    project_info.CreateFromNode drchecksnode
'    Set project_info_offset = project_info.PasteData(ActiveSheet.Range("D1"))
'    Debug.Print Cells(project_info_offset.Row, 1).Address
'
'End Sub
'
'Private Sub Test_eval1()
'    Dim root As IXMLDOMElement
'    Dim project_info As New ProjectInfo
'    Dim project_info_offset As Range
'
'    Set root = GetRootFromXML(path_concept)
'    Set a_comment = root.SelectNodes("Comments/*").Item(0)
'    Set an_eval = a_comment.SelectNodes("evaluations/*")
'
'    Debug.Print an_eval.Length
'    Debug.Print ""
'
'End Sub
'
'
'
'
'
'
'
'
'Private Sub Test_1()
'    Dim root As IXMLDOMElement
'    Set root = GetRootFromXML(path_no_comments)
'    Debug.Print "Project Info: "; root.SelectSingleNode("DrChecks").ChildNodes.Length
'    Debug.Print "Comments: "; root.SelectSingleNode("Comments").ChildNodes.Length
'
'    Debug.Print GatherNodeChildren(root.SelectSingleNode("Comments"))
'End Sub
'
'
'
'
'Private Sub Test_ProjectInfoClass()
'
'    Dim project_info As ProjectInfo
'
'    Dim project_id As String
'    Dim project_control_number As String
'    Dim project_name As String
'    Dim review_id As String
'    Dim review_name As String
'
'    project_id = "1234"
'    project_control_number = "7777"
'    project_name = "Magic Kingdom"
'    review_id = "9876"
'    review_name = "Concept submittal"
'
'    Set project_info = New ProjectInfo
'
'    'project_info.a_test = "Alice in wonderland"
'    Debug.Print project_info.ProjectID
'    Debug.Print project_info.ProjectControlNumber
'
'
'End Sub
'
'Private Sub Test_EvaluationResponse()
'    Dim root As IXMLDOMElement
'    Dim selected_eval As IXMLDOMElement
'    Dim eval_comment As New Evaluation
'    Dim sibling As Long
'    sibling = 10
'
'    Set root = GetRootFromXML(path_concept)
'    Set selected_eval = root.SelectNodes("Comments/comment/evaluations/*").Item(sibling)
'
'    eval_comment.CreateFromNode selected_eval, sibling
'    eval_comment.PrintToDebug
'End Sub
'Private Sub Test_BackCheckResponse()
'    Dim root As IXMLDOMElement
'    Dim selected_eval As IXMLDOMElement
'    Dim eval_comment As New Backcheck
'    Dim sibling As Long
'    sibling = 2
'
'    Set root = GetRootFromXML(path_concept)
'    Set selected_eval = root.SelectNodes("Comments/comment/backchecks/*").Item(sibling)
'
'    eval_comment.CreateFromNode selected_eval, sibling
'    eval_comment.PrintToDebug
'End Sub
'
'Private Sub Test_Collections()
'    Dim col As New Collection
'    Debug.Print col.Count
'End Sub
'
'
'
'
'
'
'Private Sub Test_EvaluationsClass()
'    Dim root As IXMLDOMElement
'    Dim comment_node As IXMLDOMElement
'    Dim evals As New Evaluations
'    Dim sibling As Long
'    sibling = 0
'
'    Set root = GetRootFromXML(path_concept)
'    Set comment_node = root.SelectNodes("Comments/*").Item(sibling)
'
'    evals.CreateFromNode comment_node
'    Debug.Print evals.Count
'    For Each eval In evals.List
'        eval.PrintToDebug
'        Debug.Print ""
'    Next
'    evals.Item(1).PrintToDebug
'End Sub
'
'Private Sub Test_BackchecksClass()
'    Dim root As IXMLDOMElement
'    Dim comment_node As IXMLDOMElement
'    Dim bcs As New Backchecks
'    Dim sibling As Long
'    sibling = 0
'
'    Set root = GetRootFromXML(path_concept)
'    Set comment_node = root.SelectNodes("Comments/*").Item(sibling)
'
'    bcs.CreateFromNode comment_node
'    For Each bc In bcs.List
'        bc.PrintToDebug
'        Debug.Print ""
'    Next
'    bcs.Item(1).PrintToDebug
'End Sub
'
'
'
'Private Sub Test_XPATH_exists()
'    Dim root As IXMLDOMElement
'    Set root = GetRootFromXML(path_concept)
'    Set parent_node = root.SelectNodes("Comments/*").Item(1)
'    'Debug.Print parent_node.SelectSingleNode("notanode") '''Fail'''
'    'Debug.Print parent_node.SelectSingleNode("id").Text
'
'    'Set parent_node = root.SelectSingleNode("Comments/comment[0]")
'    Set test_node = parent_node.SelectSingleNode("na")
'    'Debug.Print Not test_node Is Nothing
'
'    If Not parent_node.SelectSingleNode("sillygoose") Is Nothing Then Debug.Print "Exists" Else Debug.Print "Doesn't Exist"
'End Sub
'
'
'
'
'
'
'Private Sub Test_CommentClass()
'    Dim root As IXMLDOMElement
'    Dim a_comment_node As IXMLDOMElement
'    Dim a_comment As New Comment
'    Set root = GetRootFromXML(path_concept)
'    Set a_comment_node = root.SelectNodes("Comments/*").Item(0)
'    a_comment.CreateFromNode a_comment_node
'    a_comment.PrintToDebug
'End Sub
'
'Private Sub Test_CommentsClass()
'    Dim root As IXMLDOMElement
'    Dim all_comments As New Comments
'    Dim a_comment As New Comment
'    Set root = GetRootFromXML(path_concept)
'    all_comments.CreateFromRootElement root
'
''    Debug.Print
''    For Each a_comment In all_comments.List
''        Debug.Print a_comment.ID, a_comment.EvaluationsCount, a_comment.BackchecksCount
''    Next
''    Debug.Print
''    Debug.Print all_comments.Count
''    Debug.Print all_comments.MaxEvaluations, all_comments.MaxBackchecks
'
''    Set commentRangeRegion = all_comments.CommentsRange(ActiveSheet.Range("D7"))
''    Debug.Print commentRangeRegion.Address
''    Debug.Print
'
''    Dim i As Long
''    For Each a_comment In all_comments.List
''        Debug.Print i, a_comment.DaysOpen
''        i = i + 1
''    Next
'
''    Debug.Print all_comments.HeaderCount
''    For Each header_key In all_comments.Headers
''        Debug.Print header_key, all_comments.Headers(header_key)
''    Next
'
'    Dim a_comment_node As IXMLDOMElement
'    Set a_comment_node = root.SelectNodes("Comments/comment").Item(0)
'    a_comment.CreateFromNode a_comment_node
'
'    Dim thing As Variant
'    For Each thing In all_comments.Headers()
'        Debug.Print thing
'    Next
'
'    Debug.Print all_comments.HeaderCount
'
'End Sub
'
'Private Sub Test_CommentClass_PasteData()
'    Dim root As IXMLDOMElement
'    Dim projInfo As New ProjectInfo
'    Dim all_comments As New Comments
'    Set root = GetRootFromXML(path_concept)
'
'    projInfo.CreateFromNode root.SelectSingleNode("DrChecks")
'    projInfo.PasteData ActiveSheet.Range("D1")
'
'    all_comments.CreateFromRootElement root
'    all_comments.PasteData ActiveSheet.Range("D7")
'    Debug.Print ""
'End Sub
'
'Private Sub Test_CommentClass_PasteData_EvalRegon()
'    Dim root As IXMLDOMElement
'    Dim projInfo As New ProjectInfo
'    Dim all_comments As New Comments
'    Set root = GetRootFromXML(path_concept)
'
'    projInfo.CreateFromNode root.SelectSingleNode("DrChecks")
'    projInfo.PasteData ActiveSheet.Range("D1")
'
'    all_comments.CreateFromRootElement root
'    Debug.Print all_comments.EvaluationsHeaderRange(ActiveSheet.Range("L7")).Address
'    all_comments.EvaluationsHeaderRange(ActiveSheet.Range("L7")) = all_comments.EvaluationHeaders
'    Debug.Print ""
'End Sub
'
'
'Private Sub Test_CommentClass_GetRanges()
'    Dim root As IXMLDOMElement
'    Dim projInfo As New ProjectInfo
'    Dim all_comments As New Comments
'    Set root = GetRootFromXML(path_concept)
'
'    projInfo.CreateFromNode root.SelectSingleNode("DrChecks")
'    projInfo.PasteData ActiveSheet.Range("D1")
'
'    all_comments.CreateFromRootElement root
'    all_comments.PasteData ActiveSheet.Range("D7")
'
'    Debug.Print all_comments.CommentHeaderRange(ActiveSheet.Range("D7")).Address
'    Debug.Print all_comments.CommentRange(ActiveSheet.Range("D7")).Address
'
'    Debug.Print ""
'
'End Sub
'
'Private Sub Test_CommentClass_Responses()
'    Dim root As IXMLDOMElement
'    Dim projInfo As New ProjectInfo
'    Dim all_comments As New Comments
'    Set root = GetRootFromXML(path_concept)
'
'    projInfo.CreateFromNode root.SelectSingleNode("DrChecks")
'    projInfo.PasteData ActiveSheet.Range("D1")
'
'    all_comments.CreateFromRootElement root
'
'    'Debug.Print all_comments.MaxEvaluations, all_comments.MaxBackchecks
'    'Debug.Print all_comments.OffsetToEvaluations, all_comments.OffsetToBackchecks
'
'    eval_header = all_comments.EvaluationsHeader()
'    For i = 0 To UBound(eval_header)
'        Debug.Print eval_header(i)
'    Next
'
'    backcheck_header = all_comments.BackchecksHeader()
'    For i = 0 To UBound(backcheck_header)
'        Debug.Print backcheck_header(i)
'    Next
'End Sub
'
'
'
'Private Sub Test_CommentClass_EvalsToArray()
'    Dim root As IXMLDOMElement
'    Dim projInfo As New ProjectInfo
'    Dim all_comments As New Comments
'    Dim a_comment As New Comment
'    Dim eval_list As New Evaluations
'    Dim arr As Variant
'    Dim an_evaluation As New Evaluation
'    Set root = GetRootFromXML(path_concept)
'
'    all_comments.CreateFromRootElement root
'    'Debug.Print all_comments.MaxEvaluations
''    arr = all_comments.EvaluationsToArray()
'
''    ReDim arr(all_comments.Count - 1, 14)
''
''    i = 0
''    For Each a_comment In all_comments.List
''        n = 0
''        Debug.Print a_comment.ID
''        For Each an_evaluation In a_comment
''            thing.ID
''        Next
'''        For Each an_evaluation In a_comment.AllEvaluations
''''            arr(i, j * n + 0) = an_evaluation.Status
''''            arr(i, j * n + 1) = an_evaluation.CreatedBy
''''            arr(i, j * n + 2) = an_evaluation.CreatedOn
''''            arr(i, j * n + 3) = an_evaluation.Text
''''            arr(i, j * n + 4) = an_evaluation.Attachment
''''            n = n + 1
'''        Next
''        i = i + 1
''    Next
'
'
'    Set a_comment = all_comments.Item(1)
'    For Each an_evaluation In a_comment.EvaluationsList.List
'        Debug.Print an_evaluation.ID
'    Next
'
'End Sub
'
'
'
'
'
'
'
'
'
'
