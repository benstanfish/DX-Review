VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module to store data parsed from the DrChecks node

Dim pProjectID As String
Dim pProjectControlNumber As String
Dim pProjectName As String
Dim pReviewID As String
Dim pReviewName As String

Dim pInfoHeader As Range
Dim pInfoBody As Range

Property Get InfoHeader() As Range
    Set InfoHeader = pInfoHeader
End Property

Property Get InfoBody() As Range
    Set InfoBody = pInfoBody
End Property

Property Get ProjectID() As String
    ProjectID = pProjectID
End Property

Property Get ProjectControlNumber() As String
    ProjectControlNumber = pProjectControlNumber
End Property

Property Get ProjectName() As String
    ProjectName = pProjectName
End Property

Property Get ReviewID() As String
    ReviewID = pReviewID
End Property

Property Get ReviewName() As String
    ReviewName = pReviewName
End Property

Sub CreateFromNode(drchecksnode As IXMLDOMElement)
    pProjectID = drchecksnode.ChildNodes.Item(0).Text
    pProjectControlNumber = drchecksnode.ChildNodes.Item(1).Text
    pProjectName = drchecksnode.ChildNodes.Item(2).Text
    pReviewID = drchecksnode.ChildNodes.Item(3).Text
    pReviewName = drchecksnode.ChildNodes.Item(4).Text
End Sub

Private Sub PrintToDebug()
    Debug.Print "Project ID: "; pProjectID
    Debug.Print "Project Control Number: "; pProjectControlNumber
    Debug.Print "Project Name: "; pProjectName
    Debug.Print "Review ID: "; pReviewID
    Debug.Print "Review Name: "; pReviewName
End Sub

Function PasteData(target_cell As Range, Optional spacer_rows As Long = 2) As Range
    'Paste all the proeprties into an Excel Worksheet, return range of offset specified from end
    Application.ScreenUpdating = False
    
    target_cell.Offset(0, 0) = "Project ID"
    target_cell.Offset(1, 0) = "Project Control Number"
    target_cell.Offset(2, 0) = "Project Name"
    target_cell.Offset(3, 0) = "Review ID"
    target_cell.Offset(4, 0) = "Review Name"
    
    target_cell.Offset(0, 1) = pProjectID
    target_cell.Offset(1, 1) = pProjectControlNumber
    target_cell.Offset(2, 1) = pProjectName
    target_cell.Offset(3, 1) = pReviewID
    target_cell.Offset(4, 1) = pReviewName
    
    Set pInfoHeader = Range(target_cell, target_cell.Offset(4, 0))
    Set pInfoBody = Range(target_cell.Offset(0, 1), target_cell.Offset(4, 1))
    
    'Minor formatting of the data in this region
    Range(target_cell, target_cell.Offset(4, 0)).Font.Bold = True
    target_cell.CurrentRegion.HorizontalAlignment = xlHAlignLeft
    
    'Return the address of cell offset as specified from end cell
    Set PasteData = target_cell.Offset(4 + spacer_rows, 0)
    
    Application.ScreenUpdating = True
End Function

