Attribute VB_Name = "wMsoTabStopType"
Function MsoTabStopTypeFromString(value As String) As MsoTabStopType
    If IsNumeric(value) Then
        MsoTabStopTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTabStopLeft": MsoTabStopTypeFromString = msoTabStopLeft
        Case "msoTabStopCenter": MsoTabStopTypeFromString = msoTabStopCenter
        Case "msoTabStopRight": MsoTabStopTypeFromString = msoTabStopRight
        Case "msoTabStopDecimal": MsoTabStopTypeFromString = msoTabStopDecimal
        Case "msoTabStopMixed": MsoTabStopTypeFromString = msoTabStopMixed
    End Select
End Function

Function MsoTabStopTypeToString(value As MsoTabStopType) As String
    Select Case value
        Case msoTabStopLeft: MsoTabStopTypeToString = "msoTabStopLeft"
        Case msoTabStopCenter: MsoTabStopTypeToString = "msoTabStopCenter"
        Case msoTabStopRight: MsoTabStopTypeToString = "msoTabStopRight"
        Case msoTabStopDecimal: MsoTabStopTypeToString = "msoTabStopDecimal"
        Case msoTabStopMixed: MsoTabStopTypeToString = "msoTabStopMixed"
    End Select
End Function
