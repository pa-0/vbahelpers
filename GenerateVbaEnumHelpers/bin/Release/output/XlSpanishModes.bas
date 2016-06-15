Attribute VB_Name = "wXlSpanishModes"
Function XlSpanishModesFromString(value As String) As XlSpanishModes
    If IsNumeric(value) Then
        XlSpanishModesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSpanishTuteoOnly": XlSpanishModesFromString = xlSpanishTuteoOnly
        Case "xlSpanishTuteoAndVoseo": XlSpanishModesFromString = xlSpanishTuteoAndVoseo
        Case "xlSpanishVoseoOnly": XlSpanishModesFromString = xlSpanishVoseoOnly
    End Select
End Function

Function XlSpanishModesToString(value As XlSpanishModes) As String
    Select Case value
        Case xlSpanishTuteoOnly: XlSpanishModesToString = "xlSpanishTuteoOnly"
        Case xlSpanishTuteoAndVoseo: XlSpanishModesToString = "xlSpanishTuteoAndVoseo"
        Case xlSpanishVoseoOnly: XlSpanishModesToString = "xlSpanishVoseoOnly"
    End Select
End Function
