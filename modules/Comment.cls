VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim pID As String
Dim pSpec As String
Dim pSheet As String
Dim pDetail As String
Dim pIsCritical As String
Dim pText As String
Dim pAttach As String
Dim pDocRef As String
Dim pCreatedBy As String
Dim pCreatedOn As Date
Dim pStatus As String
Dim pDiscipline As String
Dim pDocType As String
Dim pCoordinatingDiscipline As String
Dim pSiblingSequence As Long

Dim pEvaluations As Evaluations
Dim pBackchecks As Backchecks

Dim pStatuses As Dictionary

Private Sub Class_Initialize()
    Set pEvaluations = New Evaluations
    Set pBackchecks = New Backchecks
    Set pStatuses = New Dictionary

    pStatuses.Add "concur", 0
    pStatuses.Add "for information only", 1
    pStatuses.Add "non-concur", 2
    pStatuses.Add "check and resolve", 3
    pStatuses.Add 0, "concur"
    pStatuses.Add 1, "for information only"
    pStatuses.Add 2, "non-concur"
    pStatuses.Add 3, "check and resolve"
End Sub

Property Get EvaluationsList() As Evaluations
    Set EvaluationsList = pEvaluations
End Property

Property Get BackchecksList() As Backchecks
    Set BackchecksList = pBackchecks
End Property

Property Get ID() As String
    ID = pID
End Property

Property Get Spec() As String
    Spec = pSpec
End Property

Property Get SheetRef() As String
    SheetRef = pSheet
End Property

Property Get Detail() As String
    Detail = pDetail
End Property

Property Get IsCritical() As String
    IsCritical = pIsCritical
End Property

Property Get Comment() As String
    ' Alternative wrapper function to get the comment text.
    Comment = pText
End Property

Property Get Text() As String
    Text = pText
End Property

Property Get Attachment() As String
    Attachment = pAttach
End Property

Property Get DocRef() As String
    DocRef = pDocRef
End Property

Property Get DocType() As String
    DocType = pDocType
End Property

Property Get CreatedBy() As String
    CreatedBy = pCreatedBy
End Property

Property Get CreatedOn() As Date
    CreatedOn = pCreatedOn
End Property

Property Get status() As String
    status = pStatus
End Property

Property Get Discipline() As String
    Discipline = pDiscipline
End Property

Property Get CoordinatingDiscipline() As String
    CoordinatingDiscipline = pCoordinatingDiscipline
End Property

Property Get OrderAmongSiblings() As Long
    OrderAmongSiblings = pSiblingSequence
End Property

Function EvaluationsCount() As Long
    EvaluationsCount = pEvaluations.Count
End Function

Function BackchecksCount() As Long
    BackchecksCount = pBackchecks.Count
End Function

Function DaysOpen()
    Dim last_backcheck As Backcheck
    If LCase(pStatus) = "closed" And BackchecksCount > 0 Then
        DaysOpen = CLng(DateDiff("d", pCreatedOn, Now()))
        Set last_backcheck = pBackchecks.Item(pBackchecks.Count)
        DaysOpen = DateDiff("d", pCreatedOn, last_backcheck.CreatedOn)
    Else
        ' This condition covers both "open" status comments as well as errors
        ' from the previous condition, e.g. closed/no-backcheck comments
        DaysOpen = DateDiff("d", pCreatedOn, Now())
    End If
End Function

Function HighestStatus()
    Dim an_evaluation As New Evaluation
    Dim a_backcheck As New Backcheck
    Dim highest_status As Long
    highest_status = -1
    If Not pEvaluations Is Nothing Then
        For Each an_evaluation In pEvaluations.List
            If pStatuses(LCase(an_evaluation.status)) > highest_status Then highest_status = pStatuses(LCase(an_evaluation.status))
        Next
    End If
    If Not pBackchecks Is Nothing Then
        For Each a_backcheck In pBackchecks.List
            If pStatuses(LCase(a_backcheck.status)) > highest_status Then highest_status = pStatuses(LCase(a_backcheck.status))
        Next
    End If
    HighestStatus = StrConv(pStatuses(highest_status), vbProperCase)
End Function


Sub CreateFromNode(a_node As IXMLDOMElement, Optional sibling_order As Long = 0)
    pSiblingSequence = sibling_order
    pID = a_node.SelectSingleNode("id").Text
    pSpec = a_node.SelectSingleNode("spec").Text
    pSheet = a_node.SelectSingleNode("sheet").Text
    pDetail = a_node.SelectSingleNode("detail").Text
    pIsCritical = a_node.SelectSingleNode("critical").Text
    pText = Replace(a_node.SelectSingleNode("commentText").Text, "<br />", vbCrLf)
    If Len(a_node.SelectSingleNode("attachment").Text) > 0 Then pAttach = ChrW(dxreview.DAIMARU)
    If Not a_node.SelectSingleNode("DocRef") Is Nothing Then pDocRef = a_node.SelectSingleNode("DocRef").Text
    pCreatedBy = StrConv(a_node.SelectSingleNode("createdBy").Text, vbProperCase)
    pCreatedOn = CDate(a_node.SelectSingleNode("createdOn").Text)
    pStatus = StrConv(a_node.SelectSingleNode("status").Text, vbProperCase)
    pDiscipline = a_node.SelectSingleNode("Discipline").Text
    If Not a_node.SelectSingleNode("DocType") Is Nothing Then pDocType = a_node.SelectSingleNode("DocType").Text
    If Not a_node.SelectSingleNode("CoordinatingDiscipline") Is Nothing Then pCoordinatingDiscipline = a_node.SelectSingleNode("CoordinatingDiscipline").Text
    pEvaluations.CreateFromNode a_node
    pBackchecks.CreateFromNode a_node
End Sub

Sub PrintToDebug()
    Debug.Print "ID: "; pID
    Debug.Print "Spec: "; pSpec
    Debug.Print "Sheet: "; pSheet
    Debug.Print "Detail: "; pDetail
    Debug.Print "Critical: "; pIsCritical
    Debug.Print "Text: "; pText
    Debug.Print "Attachment: "; pAttach
    Debug.Print "DocRef: "; pDocRef
    Debug.Print "CreatedBy: "; pCreatedBy
    Debug.Print "CreatedOn: "; pCreatedOn
    Debug.Print "Status: "; pStatus
    Debug.Print "Discipline: "; pDiscipline
    Debug.Print "DocType: "; pDocType
    Debug.Print "Coordinating Discipline: "; pCoordinatingDiscipline
    Debug.Print "Evaluations: "; pEvaluations.Count
    Debug.Print "Backchecks: "; pBackchecks.Count
End Sub


