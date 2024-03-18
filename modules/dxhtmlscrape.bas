Attribute VB_Name = "dxhtmlscrape"

'REF: Microsoft HTML Object Library


Public Sub WriteClassificationFromHTML()

    Application.ScreenUpdating = False

    Dim aTable As ListObject
    Dim classColumn As Range
    Dim idColumn As Range
    
    Set aTable = Workbooks(Workbooks.Count).ActiveSheet.ListObjects(1)
    Set classColumn = aTable.ListColumns("Class").DataBodyRange
    Set idColumn = aTable.ListColumns("ID").DataBodyRange
    
    Dim aDict As Dictionary
    Dim htmlPath As String
    htmlPath = GetHTMLPath()
    
    If htmlPath <> "" Then
        Set aDict = GetClassFromHTML(htmlPath)
        
        For i = 1 To idColumn.Rows.Count
            If aDict.Exists(CStr(idColumn(i))) Then classColumn(i).Value = aDict(CStr(idColumn(i)))
        Next
    End If

    Application.ScreenUpdating = True

End Sub

Public Function GetClassFromHTML(htmlPath As String)

    Dim http As Object
    Set http = New XMLHTTP60
    http.Open "GET", htmlPath, False
    http.send

    Dim html As New HTMLDocument
    html.body.innerHTML = http.responseText
    
    Dim trElems As Variant
    Dim trElem As HTMLDivElement
    Set trElems = html.getElementsByTagName("tr")
    
    Dim statusElems As Variant
    Dim statusElem As HTMLDivElement
    Set statusElems = html.getElementsByClassName("commentClassification")
    
    Dim aDict As Dictionary
    Set aDict = New Dictionary
    
    Dim i As Long
    For Each trElem In trElems
        If trElem.className = "centered" Then
            aDict.Add trElem.Children(0).innerHTML, _
                GetCommentClassification(statusElems.Item(i).innerText)
            i = i + 1
        End If
    Next

    Set GetClassFromHTML = aDict

End Function

Public Function GetCommentClassification(str As String) As String
    If Len(str) <> Len(Replace(str, "(CUI)", "")) Then
        GetCommentClassification = "CUI"
    ElseIf Len(str) <> Len(Replace(str, "(U)", "")) Then
        GetCommentClassification = "Public"
    ElseIf Len(str) <> Len(Replace(str, "(Public)", "")) Then
        GetCommentClassification = "Unclassified"
    Else
        GetCommentClassification = "None"
    End If
End Function

Public Function GetHTMLPath() As String
    ' Returns EMPTY if user cancels, otherwise returns path string
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "HTML", "*.html?"
        .Title = "Choose an HTML file"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetHTMLPath = .SelectedItems(1)
    End With
End Function
