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
Dim pDownloadDate As String
Dim pImportDate As String

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

Sub CreateFromNode(drchecksnode As IXMLDOMElement, Optional file_path As String = "")
    pProjectID = drchecksnode.ChildNodes.Item(0).Text
    pProjectControlNumber = drchecksnode.ChildNodes.Item(1).Text
    pProjectName = drchecksnode.ChildNodes.Item(2).Text
    pReviewID = drchecksnode.ChildNodes.Item(3).Text
    pReviewName = drchecksnode.ChildNodes.Item(4).Text
    pDownloadDate = ReadFileDate(file_path)
    pImportDate = Now()
End Sub

Private Sub PrintToDebug()
    Debug.Print "Project ID: "; pProjectID
    Debug.Print "Project Control Number: "; pProjectControlNumber
    Debug.Print "Project Name: "; pProjectName
    Debug.Print "Review ID: "; pReviewID
    Debug.Print "Review Name: "; pReviewName
    Debug.Print "XML Save: "; pDownloadDate
    Debug.Print "Run Date: "; pImportDate
End Sub

Function PasteData(target_cell As Range, Optional spacer_rows As Long = 2) As Range
    'Paste all the proeprties into an Excel Worksheet, return range of offset specified from end
    
    Application.ScreenUpdating = False
    
    Dim arr As Variant
    Dim arr2 As Variant
    
    arr = Array("Project ID", "Project Control Number", "Project Name", _
                "Review ID", "Review Name", "XML Save", "Run Date")
    
    arr2 = Array(pProjectID, pProjectControlNumber, pProjectName, pReviewID, _
                pReviewName, pDownloadDate, pImportDate)

    Set pInfoHeader = target_cell.Resize(UBound(arr) + 1)
    Set pInfoBody = target_cell.Offset(0, 1).Resize(UBound(arr) + 1)
    
    pInfoHeader = WorksheetFunction.Transpose(arr)
    pInfoBody = WorksheetFunction.Transpose(arr2)
    
    'Minor formatting of the data in this region
    pInfoHeader.Font.Bold = True
    pInfoHeader.CurrentRegion.HorizontalAlignment = xlHAlignLeft
    
    For i = 6 To 7
        pInfoBody(i).NumberFormat = "mm/dd/yy"
    Next

    'Return the address of cell offset as specified from end cell
    Set PasteData = target_cell.Offset(UBound(arr) + spacer_rows, 0)
    
    Application.ScreenUpdating = True
End Function

