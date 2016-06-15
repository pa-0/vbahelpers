Attribute VB_Name = "wPbInkName"
Function PbInkNameFromString(value As String) As PbInkName
    If IsNumeric(value) Then
        PbInkNameFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbInkNameCyan": PbInkNameFromString = pbInkNameCyan
        Case "pbInkNameMagenta": PbInkNameFromString = pbInkNameMagenta
        Case "pbInkNameYellow": PbInkNameFromString = pbInkNameYellow
        Case "pbInkNameBlack": PbInkNameFromString = pbInkNameBlack
        Case "pbInkNameSpot1": PbInkNameFromString = pbInkNameSpot1
        Case "pbInkNameSpot2": PbInkNameFromString = pbInkNameSpot2
        Case "pbInkNameSpot3": PbInkNameFromString = pbInkNameSpot3
        Case "pbInkNameSpot4": PbInkNameFromString = pbInkNameSpot4
        Case "pbInkNameSpot5": PbInkNameFromString = pbInkNameSpot5
        Case "pbInkNameSpot6": PbInkNameFromString = pbInkNameSpot6
        Case "pbInkNameSpot7": PbInkNameFromString = pbInkNameSpot7
        Case "pbInkNameSpot8": PbInkNameFromString = pbInkNameSpot8
        Case "pbInkNameSpot9": PbInkNameFromString = pbInkNameSpot9
        Case "pbInkNameSpot10": PbInkNameFromString = pbInkNameSpot10
        Case "pbInkNameSpot11": PbInkNameFromString = pbInkNameSpot11
        Case "pbInkNameSpot12": PbInkNameFromString = pbInkNameSpot12
    End Select
End Function

Function PbInkNameToString(value As PbInkName) As String
    Select Case value
        Case pbInkNameCyan: PbInkNameToString = "pbInkNameCyan"
        Case pbInkNameMagenta: PbInkNameToString = "pbInkNameMagenta"
        Case pbInkNameYellow: PbInkNameToString = "pbInkNameYellow"
        Case pbInkNameBlack: PbInkNameToString = "pbInkNameBlack"
        Case pbInkNameSpot1: PbInkNameToString = "pbInkNameSpot1"
        Case pbInkNameSpot2: PbInkNameToString = "pbInkNameSpot2"
        Case pbInkNameSpot3: PbInkNameToString = "pbInkNameSpot3"
        Case pbInkNameSpot4: PbInkNameToString = "pbInkNameSpot4"
        Case pbInkNameSpot5: PbInkNameToString = "pbInkNameSpot5"
        Case pbInkNameSpot6: PbInkNameToString = "pbInkNameSpot6"
        Case pbInkNameSpot7: PbInkNameToString = "pbInkNameSpot7"
        Case pbInkNameSpot8: PbInkNameToString = "pbInkNameSpot8"
        Case pbInkNameSpot9: PbInkNameToString = "pbInkNameSpot9"
        Case pbInkNameSpot10: PbInkNameToString = "pbInkNameSpot10"
        Case pbInkNameSpot11: PbInkNameToString = "pbInkNameSpot11"
        Case pbInkNameSpot12: PbInkNameToString = "pbInkNameSpot12"
    End Select
End Function
