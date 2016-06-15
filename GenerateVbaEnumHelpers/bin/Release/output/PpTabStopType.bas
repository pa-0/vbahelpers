Attribute VB_Name = "wPpTabStopType"
Function PpTabStopTypeFromString(value As String) As PpTabStopType
    If IsNumeric(value) Then
        PpTabStopTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppTabStopLeft": PpTabStopTypeFromString = ppTabStopLeft
        Case "ppTabStopCenter": PpTabStopTypeFromString = ppTabStopCenter
        Case "ppTabStopRight": PpTabStopTypeFromString = ppTabStopRight
        Case "ppTabStopDecimal": PpTabStopTypeFromString = ppTabStopDecimal
        Case "ppTabStopMixed": PpTabStopTypeFromString = ppTabStopMixed
    End Select
End Function

Function PpTabStopTypeToString(value As PpTabStopType) As String
    Select Case value
        Case ppTabStopLeft: PpTabStopTypeToString = "ppTabStopLeft"
        Case ppTabStopCenter: PpTabStopTypeToString = "ppTabStopCenter"
        Case ppTabStopRight: PpTabStopTypeToString = "ppTabStopRight"
        Case ppTabStopDecimal: PpTabStopTypeToString = "ppTabStopDecimal"
        Case ppTabStopMixed: PpTabStopTypeToString = "ppTabStopMixed"
    End Select
End Function
