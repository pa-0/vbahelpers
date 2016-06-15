Attribute VB_Name = "wXlParameterType"
Function XlParameterTypeFromString(value As String) As XlParameterType
    If IsNumeric(value) Then
        XlParameterTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPrompt": XlParameterTypeFromString = xlPrompt
        Case "xlConstant": XlParameterTypeFromString = xlConstant
        Case "xlRange": XlParameterTypeFromString = xlRange
    End Select
End Function

Function XlParameterTypeToString(value As XlParameterType) As String
    Select Case value
        Case xlPrompt: XlParameterTypeToString = "xlPrompt"
        Case xlConstant: XlParameterTypeToString = "xlConstant"
        Case xlRange: XlParameterTypeToString = "xlRange"
    End Select
End Function
