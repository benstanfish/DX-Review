VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Evaluations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Collection of evaluations for a given comment node

Dim pParentCommentID As String
Dim pItems As Collection

Property Get List()
    Set List = pItems
End Property

Private Sub Class_Initialize()
    Set pItems = New Collection
End Sub

Sub CreateFromNode(a_comment_node As IXMLDOMElement)
    Dim an_evaluation As Evaluation
    Dim evaluation_node As IXMLDOMElement
    If a_comment_node.SelectNodes("evaluations/*").Length > 0 Then
        For Each evaluation_node In a_comment_node.SelectNodes("evaluations/*")
            Set an_evaluation = New Evaluation
            an_evaluation.CreateFromNode evaluation_node
            pItems.Add an_evaluation
        Next
    End If
End Sub

Function Item(index As Long)
    Set Item = pItems.Item(index)
End Function

Function HasItems() As Boolean
    If pItems.Count <> 0 Then HasItems = True
End Function

Function Count() As Long
    Count = pItems.Count
End Function
