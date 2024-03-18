Attribute VB_Name = "dxsummary"
Private Const BORDERCOLOR = webcolors.SLATEBLUE

Public Sub GenerateStatSheet()
    Dim wb As Workbook
    Set wb = Workbooks(Workbooks.Count)
    
    Dim sht As Worksheet
    Set sht = wb.ActiveSheet
    
    If wb.Name <> ThisWorkbook.Name Then
        InsertStatisticsSheet sht, wb
    End If
End Sub

Public Sub UpdateStatSheet()
    
    Dim wb As Workbook
    Set wb = Workbooks(Workbooks.Count)
    
    Dim sht As Worksheet
    Set sht = wb.ActiveSheet
    If wb.Name <> ThisWorkbook.Name And Left(sht.Name, 4) <> "STAT" Then
        Dim statSheetName As String
        statSheetName = "STAT-" & Left(sht.Name, Len(sht.Name) - 5)
        
        Application.DisplayAlerts = False
        For i = 1 To wb.Sheets.Count
            If wb.Sheets(i).Name = statSheetName Then
                wb.Sheets(i).Delete
            End If
        Next
        Application.DisplayAlerts = True
        
        InsertStatisticsSheet sht, wb
    End If
End Sub



Public Sub InsertStatisticsSheet(srcSht As Worksheet, srcWB As Workbook)

    Dim sht As Worksheet
    Set sht = ActiveWorkbook.Sheets.Add(After:=srcSht)
    sht.Name = "STAT-" & Left(srcSht.Name, Len(srcSht.Name) - 5)
    
    Dim aTable As ListObject
    Set aTable = srcSht.ListObjects(1)
    
    ' Turn off grids
    ActiveWindow.DisplayGridlines = False
    
    
    ' Insert Sheet Title
    With sht.Range("A1")
        .Value = "Dr Checks Review Statistics"
        .Font.Size = 12
        .Font.Bold = True
    End With

    ' Insert Project Identifying Information
    Dim idRange As Range
    Set idRange = sht.Range("A3:B5")
    With idRange.Columns(1)
        .Font.Bold = True
        .ColumnWidth = 14
    End With
    idRange(1).Resize(3, 1) = WorksheetFunction.Transpose(Array("Project Name", "Review ID", "Review Name"))
    idRange(1).Offset(0, 1).Resize(3, 1) = WorksheetFunction.Transpose( _
                Array(srcSht.Cells(3, aTable.ListColumns("ID").DataBodyRange.Column).Offset(0, 1), _
                Trim(srcSht.Cells(4, aTable.ListColumns("ID").DataBodyRange.Column).Offset(0, 1)), _
                srcSht.Cells(5, aTable.ListColumns("ID").DataBodyRange.Column).Offset(0, 1)))
    idRange.HorizontalAlignment = xlHAlignLeft
    idRange(1).Offset(0, 1).Font.Bold = True

    ' Insert Overall Comment Status
    Dim overallTarget As Range
    Set overallTarget = sht.Range("A7")
    With overallTarget
        .Value = "Overall Comment Status"
        .Font.Size = 11
        .Font.Bold = True
    End With

    ' Insert By Discipline Header and Data
    Dim byDiscHeader As Range
    Set byDiscHeader = sht.Range(overallTarget.Offset(1, 0), overallTarget.Offset(1, 3))
    With byDiscHeader
        .Value = Array("By Discipline", "Open", "Closed", "Total")
        .Font.Bold = True
        .EntireRow.HorizontalAlignment = xlHAlignRight
        .Borders(xlEdgeBottom).Color = BORDERCOLOR
    End With
    byDiscHeader(1).HorizontalAlignment = xlHAlignLeft
    byDiscHeader(1).Offset(1, 0).Formula2 = "=UNIQUE(" & aTable.Name & "[Discipline])"
    
    Dim byDiscStartRow As Long, byDiscLastRow As Long
    byDiscStartRow = byDiscHeader.Row + 1
    byDiscLastRow = Range(Split(byDiscHeader(1).CurrentRegion.Address, ":")(1)).Row
    
    With sht.Range(Cells(byDiscStartRow, 2), Cells(byDiscLastRow, 3))
        .Formula2 = "=COUNTIFS(" & aTable.Name & "[Discipline],$A9," & aTable.Name & "[Status],B$8)"
    End With
    With sht.Range(Cells(byDiscStartRow, 4), Cells(byDiscLastRow, 4))
        .Formula2 = "=AGGREGATE(9,4,B9:C9)"
    End With
    Dim byDiscTotals As Range
    Set byDiscTotals = sht.Range(Cells(byDiscLastRow + 1, 1), Cells(byDiscLastRow + 1, 4))
    byDiscTotals(1) = "Grand Total"
    byDiscTotals.Borders(xlEdgeTop).Color = BORDERCOLOR
    With sht.Range(Cells(byDiscLastRow + 1, 2), Cells(byDiscLastRow + 1, 4))
        .Formula2 = "=AGGREGATE(9,4,B" & byDiscStartRow & ":B" & byDiscLastRow & ")"
    End With
    byDiscTotals.Font.Bold = True
    
    ' Insert By Author Header and Data
    Dim byAuthHeader As Range
    Set byAuthHeader = sht.Range(byDiscHeader(1).Offset(0, 5), byDiscHeader(1).Offset(0, 8))
    With byAuthHeader
        .Value = Array("By Author", "Open", "Closed", "Total")
        .Font.Bold = True
        .Borders(xlEdgeBottom).Color = BORDERCOLOR
    End With
    byAuthHeader(1).HorizontalAlignment = xlHAlignLeft
    byAuthHeader(1).Offset(1, 0).Formula2 = "=UNIQUE(" & aTable.Name & "[Author])"

    Dim byAuthStartRow As Long, byAuthLastRow As Long
    byAuthStartRow = byDiscHeader.Row + 1
    byAuthLastRow = Range(Split(byAuthHeader(1).CurrentRegion.Address, ":")(1)).Row

    With sht.Range(Cells(byAuthStartRow, 7), Cells(byAuthLastRow, 8))
        .Formula2 = "=COUNTIFS(" & aTable.Name & "[Author],$F9," & aTable.Name & "[Status],G$8)"
    End With
    With sht.Range(Cells(byAuthStartRow, 9), Cells(byAuthLastRow, 9))
        .Formula2 = "=AGGREGATE(9,4,G9:H9)"
    End With
    Dim byAuthTotals As Range
    Set byAuthTotals = sht.Range(Cells(byAuthLastRow + 1, 6), Cells(byAuthLastRow + 1, 9))
    byAuthTotals(1) = "Grand Total"
    byAuthTotals.Borders(xlEdgeTop).Color = BORDERCOLOR
    With sht.Range(Cells(byAuthLastRow + 1, 7), Cells(byAuthLastRow + 1, 9))
        .Formula2 = "=AGGREGATE(9,4,G" & byAuthStartRow & ":G" & byAuthLastRow & ")"
    End With
    byAuthTotals.Font.Bold = True
    
    ' Create By Response Status Region
    Dim byStatusHeader As Range
    Set byStatusHeader = sht.Range(byAuthHeader(1).Offset(0, 5), byAuthHeader(1).Offset(0, 8))
    With byStatusHeader
        .Resize(1, 4).Value = Array("By Response", "Open", "Closed", "Total")
        .Font.Bold = True
        .Borders(xlEdgeBottom).Color = BORDERCOLOR
    End With
    byStatusHeader(1).HorizontalAlignment = xlHAlignLeft
    byStatusHeader(1).Offset(1, 0).Resize(6, 1) = WorksheetFunction.Transpose(Array("Concur", _
        "Non-Concur", "For Information Only", "Check and Resolve", "No Response", "Grand Total"))
    Dim byStatusOpenClose As Range
    Set byStatusOpenClose = Range(byStatusHeader(1).Offset(1, 1).Address, byStatusHeader(1).Offset(4, 2).Address)
    With byStatusOpenClose
        .Formula2 = "=COUNTIFS(" & aTable.Name & "[Highest Resp.],$K" & byStatusHeader(1).Row + 1 & "," & _
            aTable.Name & "[Status],L$" & byStatusHeader(1).Row & ")"
    End With
    With Range(byStatusHeader(1).Offset(5, 1).Address, byStatusHeader(1).Offset(5, 2).Address)
        .Formula2 = "=COUNTIFS(" & aTable.Name & "[Highest Resp.],""""," & _
            aTable.Name & "[Status],L$" & byStatusHeader(1).Row & ")"
    End With
    With Range(byStatusHeader(1).Offset(1, 3).Address, byStatusHeader(1).Offset(5, 3).Address)
        .Formula2 = "=AGGREGATE(9,4,L" & byStatusHeader(1).Offset(1, 2).Row & ":M" & byStatusHeader(1).Offset(1, 3).Row & ")"
    End With
    Dim StatusTotalsRange As Range
    Set StatusTotalsRange = Range(byStatusHeader(1).Offset(6, 0).Address, byStatusHeader(1).Offset(6, 3).Address)
    With Union(StatusTotalsRange(2), StatusTotalsRange(3), StatusTotalsRange(4))
        .Formula2 = "=AGGREGATE(9,4,L" & byStatusHeader(1).Row + 1 & ":L" & StatusTotalsRange.Row - 1 & ")"
    End With
    With StatusTotalsRange
        .Font.Bold = True
        .Borders(xlEdgeTop).Color = BORDERCOLOR
    End With
    
    ' Create Open Comments by Author Region
    Dim OpenByAuthTitleRng As Range
    If byAuthLastRow > byDiscLastRow Then
        Set OpenByAuthTitleRng = Cells(byAuthLastRow + 3, 1)
    Else
        Set OpenByAuthTitleRng = Cells(byDiscLastRow + 3, 1)
    End If
    With OpenByAuthTitleRng
        .Value = "Open Comments by Author"
        .Font.Size = 11
        .Font.Bold = True
    End With
    With OpenByAuthTitleRng.Offset(1, 0)
        .Formula2 = "=TRANSPOSE(UNIQUE(" & aTable.Name & "[Author]))"
    End With
    Dim OpenByAuthHeader As Range
    Set OpenByAuthHeader = Range(OpenByAuthTitleRng.Offset(1, 0), _
        OpenByAuthTitleRng.Offset(1, OpenByAuthTitleRng.CurrentRegion.Columns.Count - 1))
    With OpenByAuthHeader
        .Font.Bold = True
        .Columns.EntireColumn.ColumnWidth = 14
        .Borders(xlEdgeBottom).Color = BORDERCOLOR
        .EntireRow.WrapText = True
        .VerticalAlignment = xlVAlignBottom
    End With
    With OpenByAuthHeader.Offset(1, 0)
        .Formula2 = "=UNIQUE(FILTER(" & aTable.Name & "[ID],(" & aTable.Name & "[Author]=A$" & OpenByAuthHeader.Row & ")*(" & aTable.Name & "[Status]=""Open""),""""))"
    End With
    Dim OpenByAuthRegion As Range
    With OpenByAuthHeader.CurrentRegion
        .HorizontalAlignment = xlHAlignLeft
        Set OpenByAuthRegion = Range(.Address)
    End With
    
    If PeopleAreAssigned(aTable) Then
    
        ' Create Open Comments by Assignee Region
        Dim OpenByAssignTitleRng As Range
        Set OpenByAssignTitleRng = Cells(Range(Split(OpenByAuthRegion.Address, ":")(1)).Row + 3, 1)
        With OpenByAssignTitleRng
            .Value = "Open Comments by Assignee"
            .Font.Size = 11
            .Font.Bold = True
        End With
        OpenByAssignTitleRng.Offset(1, 0).Formula2 = "=TRANSPOSE(UNIQUE(FILTER(" & aTable.Name & "[Assignee]," & aTable.Name & "[Assignee]<>0,"""")))"
        Dim OpenByAssignHeader As Range
        Set OpenByAssignHeader = Range(OpenByAssignTitleRng.Offset(1, 0), _
            OpenByAssignTitleRng.Offset(1, OpenByAssignTitleRng.CurrentRegion.Columns.Count - 1))
        With OpenByAssignHeader
            .Font.Bold = True
            .Columns.EntireColumn.ColumnWidth = 14
            .Borders(xlEdgeBottom).Color = BORDERCOLOR
            .EntireRow.WrapText = True
            .VerticalAlignment = xlVAlignBottom
        End With
        With OpenByAssignHeader.Offset(1, 0)
            .Formula2 = "=UNIQUE(FILTER(" & aTable.Name & "[ID],(" & aTable.Name & "[Assignee]=A$" & OpenByAssignHeader.Row & ")*(" & aTable.Name & "[Status]=""Open""),""""))"
        End With
        Dim OpenByAssignRegion As Range
        With OpenByAssignHeader.CurrentRegion
            .HorizontalAlignment = xlHAlignLeft
            Set OpenByAssignRegion = Range(.Address)
        End With
        
        ' Create Comment Status by Assignee Region
        Dim StatusByAssignTitleRng As Range
        Set StatusByAssignTitleRng = Cells(Range(Split(OpenByAssignTitleRng.CurrentRegion.Address, ":")(1)).Row + 3, 1)
        With StatusByAssignTitleRng
            .Value = "Total Comment Status by Assignee"
            .Font.Size = 11
            .Font.Bold = True
        End With
        Dim StatusByAssignHeader As Range
        Set StatusByAssignHeader = Range(StatusByAssignTitleRng(1).Offset(1, 0).Address, StatusByAssignTitleRng(1).Offset(1, 3).Address)
        With StatusByAssignHeader
            .Value = Array("Assignee", "Open", "Closed", "Total")
            .HorizontalAlignment = xlHAlignRight
            .Borders(xlEdgeBottom).Color = BORDERCOLOR
        End With
        With StatusByAssignHeader(1)
            .HorizontalAlignment = xlHAlignLeft
            .EntireRow.Font.Bold = True
        End With
        StatusByAssignHeader(1).Offset(1, 0) = "Unassigned"
        With Range(StatusByAssignHeader(1).Offset(1, 1).Address, StatusByAssignHeader(1).Offset(1, 2).Address)
            .Formula2 = "=COUNTIFS(" & aTable.Name & "[Assignee],""""," & aTable.Name & "[Status],B$" & StatusByAssignHeader(1).Offset(1, 0).Row - 1 & ")"
        End With
        With StatusByAssignHeader(1).Offset(2, 0)
            .Formula2 = "=UNIQUE(FILTER(" & aTable.Name & "[Assignee]," & aTable.Name & "[Assignee]<>0,""""))"
        End With
        Dim StatusStartRow As Long, StatusEndRow As Long
        StatusStartRow = StatusByAssignHeader(1).Offset(2, 0).Row
        StatusEndRow = Range(Split(StatusByAssignHeader(1).CurrentRegion.Address, ":")(1)).Row
        
        Dim StatusOpenClose As Range
        Set StatusOpenClose = Range(Cells(StatusStartRow, 2), Cells(StatusEndRow, 3))
        With StatusOpenClose
            .Formula2 = "=COUNTIFS(" & aTable.Name & "[Assignee],$A" & StatusOpenClose(1).Row & "," & aTable.Name & "[Status],B$" & StatusOpenClose(1).Row - 2 & ")"
        End With
        Range(Cells(StatusStartRow - 1, 4), Cells(StatusEndRow, 4)).Formula2 = "=AGGREGATE(9,4,B" & StatusOpenClose(1).Offset(-1, 0).Row & ":C" & StatusOpenClose(1).Offset(-1, 0).Row & ")"
        Dim AssigneeTotalsRow As Long
        Dim AssigneeTotalsRange As Range
        AssigneeTotalsRow = Range(Split(StatusByAssignHeader.CurrentRegion.Address, ":")(1)).Row + 1
        Set AssigneeTotalsRange = Range(Cells(AssigneeTotalsRow, 1), Cells(AssigneeTotalsRow, 4))
        With AssigneeTotalsRange
            .Formula2 = "=AGGREGATE(9,4,A" & StatusStartRow - 1 & ":A" & StatusEndRow & ")"
            .Font.Bold = True
            .Borders(xlEdgeTop).Color = BORDERCOLOR
        End With
        AssigneeTotalsRange(1) = "Grand Total"
        
    End If
    
End Sub

Public Function PeopleAreAssigned(aTable As ListObject) As Boolean
    Dim dbr As Range
    Set dbr = aTable.ListColumns("Assignee").DataBodyRange
    Dim i As Long, j As Long
    For i = 1 To dbr.Rows.Count
        If dbr.Rows(i).Value <> "" Then j = j + 1
    Next
    If j > 0 Then PeopleAreAssigned = True
End Function


















