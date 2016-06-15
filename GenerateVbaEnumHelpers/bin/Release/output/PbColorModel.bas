Attribute VB_Name = "wPbColorModel"
Function PbColorModelFromString(value As String) As PbColorModel
    If IsNumeric(value) Then
        PbColorModelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbColorModelRGB": PbColorModelFromString = pbColorModelRGB
        Case "pbColorModelCMYK": PbColorModelFromString = pbColorModelCMYK
        Case "pbColorModelGreyScale": PbColorModelFromString = pbColorModelGreyScale
        Case "pbColorModelUnknown": PbColorModelFromString = pbColorModelUnknown
    End Select
End Function

Function PbColorModelToString(value As PbColorModel) As String
    Select Case value
        Case pbColorModelRGB: PbColorModelToString = "pbColorModelRGB"
        Case pbColorModelCMYK: PbColorModelToString = "pbColorModelCMYK"
        Case pbColorModelGreyScale: PbColorModelToString = "pbColorModelGreyScale"
        Case pbColorModelUnknown: PbColorModelToString = "pbColorModelUnknown"
    End Select
End Function
