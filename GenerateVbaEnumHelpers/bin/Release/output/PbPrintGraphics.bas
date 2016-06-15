Attribute VB_Name = "wPbPrintGraphics"
Function PbPrintGraphicsFromString(value As String) As PbPrintGraphics
    If IsNumeric(value) Then
        PbPrintGraphicsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPrintHighResolution": PbPrintGraphicsFromString = pbPrintHighResolution
        Case "pbPrintLowResolution": PbPrintGraphicsFromString = pbPrintLowResolution
        Case "pbPrintNoGraphics": PbPrintGraphicsFromString = pbPrintNoGraphics
    End Select
End Function

Function PbPrintGraphicsToString(value As PbPrintGraphics) As String
    Select Case value
        Case pbPrintHighResolution: PbPrintGraphicsToString = "pbPrintHighResolution"
        Case pbPrintLowResolution: PbPrintGraphicsToString = "pbPrintLowResolution"
        Case pbPrintNoGraphics: PbPrintGraphicsToString = "pbPrintNoGraphics"
    End Select
End Function
