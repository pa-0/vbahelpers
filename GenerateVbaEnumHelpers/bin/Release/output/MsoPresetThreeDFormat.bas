Attribute VB_Name = "wMsoPresetThreeDFormat"
Function MsoPresetThreeDFormatFromString(value As String) As MsoPresetThreeDFormat
    If IsNumeric(value) Then
        MsoPresetThreeDFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoThreeD1": MsoPresetThreeDFormatFromString = msoThreeD1
        Case "msoThreeD2": MsoPresetThreeDFormatFromString = msoThreeD2
        Case "msoThreeD3": MsoPresetThreeDFormatFromString = msoThreeD3
        Case "msoThreeD4": MsoPresetThreeDFormatFromString = msoThreeD4
        Case "msoThreeD5": MsoPresetThreeDFormatFromString = msoThreeD5
        Case "msoThreeD6": MsoPresetThreeDFormatFromString = msoThreeD6
        Case "msoThreeD7": MsoPresetThreeDFormatFromString = msoThreeD7
        Case "msoThreeD8": MsoPresetThreeDFormatFromString = msoThreeD8
        Case "msoThreeD9": MsoPresetThreeDFormatFromString = msoThreeD9
        Case "msoThreeD10": MsoPresetThreeDFormatFromString = msoThreeD10
        Case "msoThreeD11": MsoPresetThreeDFormatFromString = msoThreeD11
        Case "msoThreeD12": MsoPresetThreeDFormatFromString = msoThreeD12
        Case "msoThreeD13": MsoPresetThreeDFormatFromString = msoThreeD13
        Case "msoThreeD14": MsoPresetThreeDFormatFromString = msoThreeD14
        Case "msoThreeD15": MsoPresetThreeDFormatFromString = msoThreeD15
        Case "msoThreeD16": MsoPresetThreeDFormatFromString = msoThreeD16
        Case "msoThreeD17": MsoPresetThreeDFormatFromString = msoThreeD17
        Case "msoThreeD18": MsoPresetThreeDFormatFromString = msoThreeD18
        Case "msoThreeD19": MsoPresetThreeDFormatFromString = msoThreeD19
        Case "msoThreeD20": MsoPresetThreeDFormatFromString = msoThreeD20
        Case "msoPresetThreeDFormatMixed": MsoPresetThreeDFormatFromString = msoPresetThreeDFormatMixed
    End Select
End Function

Function MsoPresetThreeDFormatToString(value As MsoPresetThreeDFormat) As String
    Select Case value
        Case msoThreeD1: MsoPresetThreeDFormatToString = "msoThreeD1"
        Case msoThreeD2: MsoPresetThreeDFormatToString = "msoThreeD2"
        Case msoThreeD3: MsoPresetThreeDFormatToString = "msoThreeD3"
        Case msoThreeD4: MsoPresetThreeDFormatToString = "msoThreeD4"
        Case msoThreeD5: MsoPresetThreeDFormatToString = "msoThreeD5"
        Case msoThreeD6: MsoPresetThreeDFormatToString = "msoThreeD6"
        Case msoThreeD7: MsoPresetThreeDFormatToString = "msoThreeD7"
        Case msoThreeD8: MsoPresetThreeDFormatToString = "msoThreeD8"
        Case msoThreeD9: MsoPresetThreeDFormatToString = "msoThreeD9"
        Case msoThreeD10: MsoPresetThreeDFormatToString = "msoThreeD10"
        Case msoThreeD11: MsoPresetThreeDFormatToString = "msoThreeD11"
        Case msoThreeD12: MsoPresetThreeDFormatToString = "msoThreeD12"
        Case msoThreeD13: MsoPresetThreeDFormatToString = "msoThreeD13"
        Case msoThreeD14: MsoPresetThreeDFormatToString = "msoThreeD14"
        Case msoThreeD15: MsoPresetThreeDFormatToString = "msoThreeD15"
        Case msoThreeD16: MsoPresetThreeDFormatToString = "msoThreeD16"
        Case msoThreeD17: MsoPresetThreeDFormatToString = "msoThreeD17"
        Case msoThreeD18: MsoPresetThreeDFormatToString = "msoThreeD18"
        Case msoThreeD19: MsoPresetThreeDFormatToString = "msoThreeD19"
        Case msoThreeD20: MsoPresetThreeDFormatToString = "msoThreeD20"
        Case msoPresetThreeDFormatMixed: MsoPresetThreeDFormatToString = "msoPresetThreeDFormatMixed"
    End Select
End Function
