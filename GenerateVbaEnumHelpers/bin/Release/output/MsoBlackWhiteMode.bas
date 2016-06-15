Attribute VB_Name = "wMsoBlackWhiteMode"
Function MsoBlackWhiteModeFromString(value As String) As MsoBlackWhiteMode
    If IsNumeric(value) Then
        MsoBlackWhiteModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBlackWhiteAutomatic": MsoBlackWhiteModeFromString = msoBlackWhiteAutomatic
        Case "msoBlackWhiteGrayScale": MsoBlackWhiteModeFromString = msoBlackWhiteGrayScale
        Case "msoBlackWhiteLightGrayScale": MsoBlackWhiteModeFromString = msoBlackWhiteLightGrayScale
        Case "msoBlackWhiteInverseGrayScale": MsoBlackWhiteModeFromString = msoBlackWhiteInverseGrayScale
        Case "msoBlackWhiteGrayOutline": MsoBlackWhiteModeFromString = msoBlackWhiteGrayOutline
        Case "msoBlackWhiteBlackTextAndLine": MsoBlackWhiteModeFromString = msoBlackWhiteBlackTextAndLine
        Case "msoBlackWhiteHighContrast": MsoBlackWhiteModeFromString = msoBlackWhiteHighContrast
        Case "msoBlackWhiteBlack": MsoBlackWhiteModeFromString = msoBlackWhiteBlack
        Case "msoBlackWhiteWhite": MsoBlackWhiteModeFromString = msoBlackWhiteWhite
        Case "msoBlackWhiteDontShow": MsoBlackWhiteModeFromString = msoBlackWhiteDontShow
        Case "msoBlackWhiteMixed": MsoBlackWhiteModeFromString = msoBlackWhiteMixed
    End Select
End Function

Function MsoBlackWhiteModeToString(value As MsoBlackWhiteMode) As String
    Select Case value
        Case msoBlackWhiteAutomatic: MsoBlackWhiteModeToString = "msoBlackWhiteAutomatic"
        Case msoBlackWhiteGrayScale: MsoBlackWhiteModeToString = "msoBlackWhiteGrayScale"
        Case msoBlackWhiteLightGrayScale: MsoBlackWhiteModeToString = "msoBlackWhiteLightGrayScale"
        Case msoBlackWhiteInverseGrayScale: MsoBlackWhiteModeToString = "msoBlackWhiteInverseGrayScale"
        Case msoBlackWhiteGrayOutline: MsoBlackWhiteModeToString = "msoBlackWhiteGrayOutline"
        Case msoBlackWhiteBlackTextAndLine: MsoBlackWhiteModeToString = "msoBlackWhiteBlackTextAndLine"
        Case msoBlackWhiteHighContrast: MsoBlackWhiteModeToString = "msoBlackWhiteHighContrast"
        Case msoBlackWhiteBlack: MsoBlackWhiteModeToString = "msoBlackWhiteBlack"
        Case msoBlackWhiteWhite: MsoBlackWhiteModeToString = "msoBlackWhiteWhite"
        Case msoBlackWhiteDontShow: MsoBlackWhiteModeToString = "msoBlackWhiteDontShow"
        Case msoBlackWhiteMixed: MsoBlackWhiteModeToString = "msoBlackWhiteMixed"
    End Select
End Function
