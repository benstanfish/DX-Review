Attribute VB_Name = "testing2"
'' Module requires references to:
''    Microsoft XML, v6.0 (msxml6.dll) - XML parsing functions
''    Microsoft Scripting Runtime (scrrun.dll) - Dictionaries
'
'Const path_concept = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\concept.xml"
'Const path_no_responses = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\no responses.xml"
'Const path_books = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\books.xml"
'Const path_middle = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\middle.xml"
'Const path_no_bcs = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\small - no bc.xml"
'Const path_final = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\final.xml"
'Const path_vaal1 = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\vaal1.xml"
'Const path_vaal2 = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\vaal2.xml"
'
'Const path_test1 = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\test1.xml"
'Const path_test2 = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\test2.xml"
'Const path_test4 = "C:\Users\benst\Documents\_0 Workspace\py\DX-Review\dev\test4.xml"
'
'Public Sub ImportFromXML()
'    Dim root As IXMLDOMElement
'    Dim projInfo As New ProjectInfo
'    Dim all_comments As New Comments
'    Dim user_notes As New UserNotes
'    Set root = GetRootFromXML(path_no_responses)
'
'    ActiveSheet.Cells.Clear
'
'    projInfo.CreateFromNode root.SelectSingleNode("DrChecks")
'    projInfo.PasteData ActiveSheet.Range("E1")
'
'    all_comments.CreateFromRootElement root
'    all_comments.PasteData ActiveSheet.Range("E7")
'
'    user_notes.PasteData ActiveSheet.Range("A7"), all_comments.Count
'
''    Debug.Print "PI H: "; projInfo.InfoHeader.Address
''    Debug.Print "PI B: "; projInfo.InfoBody.Address
'
'
''    Debug.Print all_comments.Count
''    Debug.Print all_comments.MaxEvaluations, all_comments.MaxBackchecks
''    Debug.Print all_comments.ResponseHeaderFieldCount
''    Debug.Print ""
''    Debug.Print "C H: "; all_comments.CommentsHeader.Address
''    Debug.Print "C B: "; all_comments.CommentsBody.Address
''    Debug.Print "Eval H: "; all_comments.EvaluationsHeader.Address
''    Debug.Print "Eval B: "; all_comments.EvaluationsBody.Address
''    Debug.Print "BC H: "; all_comments.BackchecksHeader.Address
''    Debug.Print "BC B: "; all_comments.BackchecksBody.Address
'
''    Debug.Print "AR H: "; all_comments.AllResponseHeader.Address
''    Debug.Print "AR B: "; all_comments.AllResponseBody.Address
'
'    all_comments.ApplyFormats
'    user_notes.ApplyFormats
'
'End Sub
'
'
'
'
'
'
'
