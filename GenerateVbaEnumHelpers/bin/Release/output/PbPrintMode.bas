Attribute VB_Name = "wPbPrintMode"
Function PbPrintModeFromString(value As String) As PbPrintMode
    If IsNumeric(value) Then
        PbPrintModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPrintModeCompositeRGB": PbPrintModeFromString = pbPrintModeCompositeRGB
        Case "pbPrintModeSeparations": PbPrintModeFromString = pbPrintModeSeparations
        Case "pbPrintModeCompositeCMYK": PbPrintModeFromString = pbPrintModeCompositeCMYK
        Case "pbPrintModeCompositeGrayscale": PbPrintModeFromString = pbPrintModeCompositeGrayscale
    End Select
End Function

Function PbPrintModeToString(value As PbPrintMode) As String
    Select Case value
        Case pbPrintModeCompositeRGB: PbPrintModeToString = "pbPrintModeCompositeRGB"
        Case pbPrintModeSeparations: PbPrintModeToString = "pbPrintModeSeparations"
        Case pbPrintModeCompositeCMYK: PbPrintModeToString = "pbPrintModeCompositeCMYK"
        Case pbPrintModeCompositeGrayscale: PbPrintModeToString = "pbPrintModeCompositeGrayscale"
    End Select
End Function
