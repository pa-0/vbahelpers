Attribute VB_Name = "wMsoDiagramNodeType"
Function MsoDiagramNodeTypeFromString(value As String) As MsoDiagramNodeType
    If IsNumeric(value) Then
        MsoDiagramNodeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoDiagramNode": MsoDiagramNodeTypeFromString = msoDiagramNode
        Case "msoDiagramAssistant": MsoDiagramNodeTypeFromString = msoDiagramAssistant
    End Select
End Function

Function MsoDiagramNodeTypeToString(value As MsoDiagramNodeType) As String
    Select Case value
        Case msoDiagramNode: MsoDiagramNodeTypeToString = "msoDiagramNode"
        Case msoDiagramAssistant: MsoDiagramNodeTypeToString = "msoDiagramAssistant"
    End Select
End Function
