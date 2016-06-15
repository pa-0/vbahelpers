Attribute VB_Name = "wXlXLMMacroType"
Function XlXLMMacroTypeFromString(value As String) As XlXLMMacroType
    If IsNumeric(value) Then
        XlXLMMacroTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFunction": XlXLMMacroTypeFromString = xlFunction
        Case "xlCommand": XlXLMMacroTypeFromString = xlCommand
        Case "xlNotXLM": XlXLMMacroTypeFromString = xlNotXLM
    End Select
End Function

Function XlXLMMacroTypeToString(value As XlXLMMacroType) As String
    Select Case value
        Case xlFunction: XlXLMMacroTypeToString = "xlFunction"
        Case xlCommand: XlXLMMacroTypeToString = "xlCommand"
        Case xlNotXLM: XlXLMMacroTypeToString = "xlNotXLM"
    End Select
End Function
