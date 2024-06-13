VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim pHeaderRange As Range
Dim pBodyRange As Range

Dim pHeaderFields As Variant

Private Sub Class_Initialize()
    pHeaderFields = Array("No.", "Notes", "Action Items", "Assignee", _
                          "Proposed Response", "Proposed Status", "State")
End Sub

Sub PasteData(target_cell As Range, comment_count As Long)
    Application.ScreenUpdating = False
    Dim arr As Variant, i As Long
    ReDim arr(UBound(pHeaderFields) - 1)
    For i = 0 To UBound(arr)
        arr(i) = pHeaderFields(i)
    Next
    
    Set pHeaderRange = Range(target_cell, target_cell.Offset(0, UBound(arr) + 1))
    Set pBodyRange = Range(target_cell.Offset(1, 0), target_cell.Offset(comment_count, UBound(arr) + 1))
    
    For i = 1 To pBodyRange.Rows.Count
        pBodyRange(1).Offset(i - 1, 0) = i
    Next
    
    pHeaderRange = pHeaderFields
    Application.ScreenUpdating = True
End Sub

Sub ApplyFormats()
    Application.ScreenUpdating = False
    With pHeaderRange
        .Interior.Color = webcolors.LEMONCHIFFON
        .Font.Color = BLACK
        
        .Borders(xlEdgeBottom).LineStyle = xlNone
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = DARKGOLDENROD
        End With
        
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignBottom
    End With
    
    With pBodyRange
        .Interior.Color = LIGHTGOLDENRODYELLOW
        .Font.Color = SADDLEBROWN
        
        .Borders(xlEdgeBottom).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = KHAKI
        End With
        .Font.Size = 8
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With

    pHeaderRange(1).ColumnWidth = 5
    For i = 2 To 5
        With pHeaderRange(i)
            .ColumnWidth = 20
            .WrapText = True
        End With
    Next
    With pHeaderRange(6)
        .ColumnWidth = 12
        .EntireColumn.WrapText = True
    End With
    pHeaderRange(7).ColumnWidth = 8
    
    Range(pHeaderRange(2), pHeaderRange(pHeaderRange.Count)).Group
    ActiveSheet.Outline.ShowLevels ColumnLevels:=1
    Application.ScreenUpdating = True
End Sub
