Attribute VB_Name = "dxreviewv2_stash"

Function get_nodetype_from_enum(node_type_number)
   
    Dim arr As Variant
    arr = Array("NODE_ELEMENT", _
                "NODE_ATTRIBUTE", _
                "NODE_TEXT", _
                "NODE_CDATA_SECTION", _
                "NODE_ENTITY_REFERENCE", _
                "NODE_ENTITY", _
                "NODE_PROCESSING_INSTRUCTION", _
                "NODE_COMMENT", _
                "NODE_DOCUMENT", _
                "NODE_DOCUMENT_TYPE", _
                "NODE_DOCUMENT_FRAGMENT", _
                "NODE_NOTATION")
    get_nodetype_from_enum = arr(node_type_number - 1)
    
End Function

Function count_evaluations(root_element As IXMLDOMElement) as Long
    'Returns a count of all <evaulation*> elements found in the XML file.
    Dim evaluations As IXMLDOMSelection
    Dim i, count As Long
    Set evaluations = root_element.selectNodes("Comments/comment/evaluations")
    For i = 0 To evaluations.Length - 1
        count = count + evaluations(i).ChildNodes.Length
    Next
    count_evaluations = count
End Function

Function count_backchecks(root_element As IXMLDOMElement) as Long
    'Returns a count of all <backcheck*> elements found in the XML file.
    Dim backchecks As IXMLDOMSelection
    Dim i, count As Long
    Set backchecks = root_element.selectNodes("Comments/comment/backchecks")
    For i = 0 To backchecks.Length - 1
        count = count + backchecks(i).ChildNodes.Length
    Next
    count_backchecks = count
End Function

Function get_max_evaluations(root_element As IXMLDOMElement) as Long
    Dim evaluations As IXMLDOMSelection
    Dim i, max_count As Long
    Set evaluations = root_element.selectNodes("Comments/comment/evaluations")
    For i = 0 To evaluations.Length - 1
        If max_count < evaluations(i).ChildNodes.Length Then
            max_count = evaluations(i).ChildNodes.Length
        End If
    Next
    get_max_evaluations = max_count
End Function

Function get_max_backchecks(root_element As IXMLDOMElement)
    ' COMPLETED
    Dim backchecks As IXMLDOMSelection
    Dim max_count As Long
    Set backchecks = root_element.selectNodes("Comments/comment/backchecks")
    For i = 0 To backchecks.Length - 1
        If max_count < backchecks(i).ChildNodes.Length Then
            max_count = backchecks(i).ChildNodes.Length
        End If
    Next
    get_max_backchecks = max_count
End Function

Function count_comment_evaluations(ByVal a_node As IXMLDOMNode) As Long
    ' COMPLETED
    Set evaulations = a_node.selectNodes("evaluations")
    count_comment_evaluations = evaulations.Item(0).ChildNodes.Length
End Function

Function count_comment_backchecks(ByVal a_node As IXMLDOMNode) As Long
    ' COMPLETED
    Set backchecks = a_node.selectNodes("backchecks")
    count_comment_backchecks = backchecks.Item(0).ChildNodes.Length
End Function

'This method replaces count_comment_evaluations() and count_comment_backchecks()
Function count_children_by_type(element_as_XPATH as String, ByVal parent_node As IXMLDOMNode)
    'Used to count the <evaluation*> or <backcheck*> children of the <evaluations> or 
    '<backchecks> nodes of a parent_node, typically with XPaths = "evaluations" or "backchecks"
    count_children_by_type = parent_node.selectNodes(element_as_XPATH).Item(0).ChildNodes.Length
End Function