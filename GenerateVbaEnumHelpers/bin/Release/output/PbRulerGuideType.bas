Attribute VB_Name = "wPbRulerGuideType"
Function PbRulerGuideTypeFromString(value As String) As PbRulerGuideType
    If IsNumeric(value) Then
        PbRulerGuideTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbRulerGuideTypeVertical": PbRulerGuideTypeFromString = pbRulerGuideTypeVertical
        Case "pbRulerGuideTypeHorizontal": PbRulerGuideTypeFromString = pbRulerGuideTypeHorizontal
    End Select
End Function

Function PbRulerGuideTypeToString(value As PbRulerGuideType) As String
    Select Case value
        Case pbRulerGuideTypeVertical: PbRulerGuideTypeToString = "pbRulerGuideTypeVertical"
        Case pbRulerGuideTypeHorizontal: PbRulerGuideTypeToString = "pbRulerGuideTypeHorizontal"
    End Select
End Function
