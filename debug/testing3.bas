Attribute VB_Name = "testing3"
'
'
'
'Const search_pattern_1 = "\(\d+?\)"
'Const search_pattern = "(?:\()\d+?(?:\))"
'
'Private Sub test_find_index()
'
'    Dim file_name As String
'
'    file_name = "here is a file name(10)"
'    file_name_1 = "here is a file name"
'
'    Dim regex As New RegExp
'    Dim a_match As Match
'    Dim match_collection As MatchCollection
'
'    With regex
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'        .Pattern = search_pattern_1
'    End With
'
'    Set match_collection = regex.Execute(file_name)
'    Debug.Print match_collection(0)
'    For Each a_match In match_collection
'        regex.Pattern = "\(" & a_match & "\)"
'        Debug.Print regex.Replace(file_name, "")
'
''        current_value = CLng(Replace(Replace(a_match.value, "(", ""), ")", ""))
''        next_value$ = "(" & current_value + 1 & ")"
''        Debug.Print next_value
''
''        Debug.Print regex.Replace(file_name, next_value)
'    Next
'
''    Set match_collection = regex.Execute(file_name_1)
''    current_value = CLng(Replace(Replace(match_collection(0).value, "(", ""), ")", ""))
''    next_value$ = "(" & current_value + 1 & ")"
''    Debug.Print next_value
''
''    Debug.Print regex.Replace(file_name, next_value)
'
'End Sub
'
'Private Sub rename_sheet()
'
'    Dim wb As Workbook
'    Dim ws As Worksheet
'    Dim ws_Name As String
'
'    Set wb = ActiveWorkbook
'    ws_Name = "Sheet9"
'
'    Dim regex As New RegExp
'    Dim a_match As Match
'    Dim match_collection As MatchCollection
'
'    Dim index_pattern As String
'    index_pattern = "\(\d+?\)"
'
'    With regex
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'        .Pattern = index_pattern
'    End With
'
'    Set match_collection = regex.Execute(ws_Name)
'    If match_collection.Count > 0 Then
'        base_name = match_collection(0)
'    Else
'        base_name = ws_Name
'    End If
'    Debug.Print base_name
'
'    For Each ws In wb.Sheets
'        Debug.Print ws.Name = ws_Name
'    Next
'
'End Sub
'
