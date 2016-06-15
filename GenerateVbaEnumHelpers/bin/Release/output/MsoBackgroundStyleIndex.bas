Attribute VB_Name = "wMsoBackgroundStyleIndex"
Function MsoBackgroundStyleIndexFromString(value As String) As MsoBackgroundStyleIndex
    If IsNumeric(value) Then
        MsoBackgroundStyleIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBackgroundStyleNotAPreset": MsoBackgroundStyleIndexFromString = msoBackgroundStyleNotAPreset
        Case "msoBackgroundStylePreset1": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset1
        Case "msoBackgroundStylePreset2": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset2
        Case "msoBackgroundStylePreset3": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset3
        Case "msoBackgroundStylePreset4": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset4
        Case "msoBackgroundStylePreset5": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset5
        Case "msoBackgroundStylePreset6": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset6
        Case "msoBackgroundStylePreset7": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset7
        Case "msoBackgroundStylePreset8": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset8
        Case "msoBackgroundStylePreset9": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset9
        Case "msoBackgroundStylePreset10": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset10
        Case "msoBackgroundStylePreset11": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset11
        Case "msoBackgroundStylePreset12": MsoBackgroundStyleIndexFromString = msoBackgroundStylePreset12
        Case "msoBackgroundStyleMixed": MsoBackgroundStyleIndexFromString = msoBackgroundStyleMixed
    End Select
End Function

Function MsoBackgroundStyleIndexToString(value As MsoBackgroundStyleIndex) As String
    Select Case value
        Case msoBackgroundStyleNotAPreset: MsoBackgroundStyleIndexToString = "msoBackgroundStyleNotAPreset"
        Case msoBackgroundStylePreset1: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset1"
        Case msoBackgroundStylePreset2: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset2"
        Case msoBackgroundStylePreset3: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset3"
        Case msoBackgroundStylePreset4: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset4"
        Case msoBackgroundStylePreset5: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset5"
        Case msoBackgroundStylePreset6: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset6"
        Case msoBackgroundStylePreset7: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset7"
        Case msoBackgroundStylePreset8: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset8"
        Case msoBackgroundStylePreset9: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset9"
        Case msoBackgroundStylePreset10: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset10"
        Case msoBackgroundStylePreset11: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset11"
        Case msoBackgroundStylePreset12: MsoBackgroundStyleIndexToString = "msoBackgroundStylePreset12"
        Case msoBackgroundStyleMixed: MsoBackgroundStyleIndexToString = "msoBackgroundStyleMixed"
    End Select
End Function
