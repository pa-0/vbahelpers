Attribute VB_Name = "wMsoSmartArtNodeType"
Function MsoSmartArtNodeTypeFromString(value As String) As MsoSmartArtNodeType
    If IsNumeric(value) Then
        MsoSmartArtNodeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSmartArtNodeTypeDefault": MsoSmartArtNodeTypeFromString = msoSmartArtNodeTypeDefault
        Case "msoSmartArtNodeTypeAssistant": MsoSmartArtNodeTypeFromString = msoSmartArtNodeTypeAssistant
    End Select
End Function

Function MsoSmartArtNodeTypeToString(value As MsoSmartArtNodeType) As String
    Select Case value
        Case msoSmartArtNodeTypeDefault: MsoSmartArtNodeTypeToString = "msoSmartArtNodeTypeDefault"
        Case msoSmartArtNodeTypeAssistant: MsoSmartArtNodeTypeToString = "msoSmartArtNodeTypeAssistant"
    End Select
End Function
