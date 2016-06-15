Attribute VB_Name = "wWdXMLNodeType"
Function WdXMLNodeTypeFromString(value As String) As WdXMLNodeType
    If IsNumeric(value) Then
        WdXMLNodeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdXMLNodeElement": WdXMLNodeTypeFromString = wdXMLNodeElement
        Case "wdXMLNodeAttribute": WdXMLNodeTypeFromString = wdXMLNodeAttribute
    End Select
End Function

Function WdXMLNodeTypeToString(value As WdXMLNodeType) As String
    Select Case value
        Case wdXMLNodeElement: WdXMLNodeTypeToString = "wdXMLNodeElement"
        Case wdXMLNodeAttribute: WdXMLNodeTypeToString = "wdXMLNodeAttribute"
    End Select
End Function
