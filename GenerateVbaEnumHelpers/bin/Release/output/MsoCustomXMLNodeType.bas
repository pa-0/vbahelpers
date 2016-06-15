Attribute VB_Name = "wMsoCustomXMLNodeType"
Function MsoCustomXMLNodeTypeFromString(value As String) As MsoCustomXMLNodeType
    If IsNumeric(value) Then
        MsoCustomXMLNodeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCustomXMLNodeElement": MsoCustomXMLNodeTypeFromString = msoCustomXMLNodeElement
        Case "msoCustomXMLNodeAttribute": MsoCustomXMLNodeTypeFromString = msoCustomXMLNodeAttribute
        Case "msoCustomXMLNodeText": MsoCustomXMLNodeTypeFromString = msoCustomXMLNodeText
        Case "msoCustomXMLNodeCData": MsoCustomXMLNodeTypeFromString = msoCustomXMLNodeCData
        Case "msoCustomXMLNodeProcessingInstruction": MsoCustomXMLNodeTypeFromString = msoCustomXMLNodeProcessingInstruction
        Case "msoCustomXMLNodeComment": MsoCustomXMLNodeTypeFromString = msoCustomXMLNodeComment
        Case "msoCustomXMLNodeDocument": MsoCustomXMLNodeTypeFromString = msoCustomXMLNodeDocument
    End Select
End Function

Function MsoCustomXMLNodeTypeToString(value As MsoCustomXMLNodeType) As String
    Select Case value
        Case msoCustomXMLNodeElement: MsoCustomXMLNodeTypeToString = "msoCustomXMLNodeElement"
        Case msoCustomXMLNodeAttribute: MsoCustomXMLNodeTypeToString = "msoCustomXMLNodeAttribute"
        Case msoCustomXMLNodeText: MsoCustomXMLNodeTypeToString = "msoCustomXMLNodeText"
        Case msoCustomXMLNodeCData: MsoCustomXMLNodeTypeToString = "msoCustomXMLNodeCData"
        Case msoCustomXMLNodeProcessingInstruction: MsoCustomXMLNodeTypeToString = "msoCustomXMLNodeProcessingInstruction"
        Case msoCustomXMLNodeComment: MsoCustomXMLNodeTypeToString = "msoCustomXMLNodeComment"
        Case msoCustomXMLNodeDocument: MsoCustomXMLNodeTypeToString = "msoCustomXMLNodeDocument"
    End Select
End Function
