Attribute VB_Name = "wPpPrintColorType"
Function PpPrintColorTypeFromString(value As String) As PpPrintColorType
    If IsNumeric(value) Then
        PpPrintColorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPrintColor": PpPrintColorTypeFromString = ppPrintColor
        Case "ppPrintBlackAndWhite": PpPrintColorTypeFromString = ppPrintBlackAndWhite
        Case "ppPrintPureBlackAndWhite": PpPrintColorTypeFromString = ppPrintPureBlackAndWhite
    End Select
End Function

Function PpPrintColorTypeToString(value As PpPrintColorType) As String
    Select Case value
        Case ppPrintColor: PpPrintColorTypeToString = "ppPrintColor"
        Case ppPrintBlackAndWhite: PpPrintColorTypeToString = "ppPrintBlackAndWhite"
        Case ppPrintPureBlackAndWhite: PpPrintColorTypeToString = "ppPrintPureBlackAndWhite"
    End Select
End Function
