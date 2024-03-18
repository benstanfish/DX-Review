VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Backcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module to store data parsed from an evaluation or backcheck node.

Dim pID As String
Dim pParentCommentID As String
Dim pRelatedEvaluationID As String
Dim pStatus As String
Dim pText As String
Dim pAttach As String
Dim pCreatedBy As String
Dim pCreatedOn As Date
Dim pSiblingSequence As Long

Property Get ParentID() As String
    ParentID = pParentCommentID
End Property

Property Get RelatedEvaluationID() As String
    RelatedEvaluationID = pRelatedEvaluationID
End Property

Property Get ID() As String
    ID = pID
End Property

Property Get status() As String
    status = pStatus
End Property

Property Get Text() As String
    Text = pText
End Property

Property Get Attachment() As String
    Attachment = pAttach
End Property

Property Get CreatedBy() As String
    CreatedBy = pCreatedBy
End Property

Property Get CreatedOn() As Date
    CreatedOn = pCreatedOn
End Property

Property Get OrderAmongSiblings() As Long
    OrderAmongSiblings = pSiblingSequence
End Property

Sub CreateFromNode(a_node As IXMLDOMElement, Optional sibling_order As Long = 0)
    ' a_node must be the backcheck node, otherwise an error results
    pSiblingSequence = sibling_order
    pID = a_node.SelectSingleNode("id").Text
    pParentCommentID = a_node.SelectSingleNode("comment").Text
    pRelatedEvaluationID = a_node.SelectSingleNode("evaluation").Text
    pStatus = a_node.SelectSingleNode("status").Text
    pText = Replace(a_node.SelectSingleNode("backcheckText").Text, "<br />", vbCrLf)
    If Len(a_node.SelectSingleNode("attachment").Text) > 0 Then pAttach = ChrW(dxreview.DAIMARU)
    pCreatedBy = StrConv(a_node.SelectSingleNode("createdBy").Text, vbProperCase)
    pCreatedOn = CDate(a_node.SelectSingleNode("createdOn").Text)
End Sub

Sub PrintToDebug()
    Debug.Print "ID: "; pID
    Debug.Print "Parent Comment ID: "; pParentCommentID
    Debug.Print "Related Evaluation ID: "; pRelatedEvaluationID
    Debug.Print "Status: "; pStatus
    Debug.Print "Comment Text: "; pText
    Debug.Print "Attachment: "; pAttach
    Debug.Print "Created By: "; pCreatedBy
    Debug.Print "Created On: "; pCreatedOn
    Debug.Print "Order Among Siblings: "; pSiblingSequence
End Sub

