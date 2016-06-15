Attribute VB_Name = "wXlPhoneticAlignment"
Function XlPhoneticAlignmentFromString(value As String) As XlPhoneticAlignment
    If IsNumeric(value) Then
        XlPhoneticAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPhoneticAlignNoControl": XlPhoneticAlignmentFromString = xlPhoneticAlignNoControl
        Case "xlPhoneticAlignLeft": XlPhoneticAlignmentFromString = xlPhoneticAlignLeft
        Case "xlPhoneticAlignCenter": XlPhoneticAlignmentFromString = xlPhoneticAlignCenter
        Case "xlPhoneticAlignDistributed": XlPhoneticAlignmentFromString = xlPhoneticAlignDistributed
    End Select
End Function

Function XlPhoneticAlignmentToString(value As XlPhoneticAlignment) As String
    Select Case value
        Case xlPhoneticAlignNoControl: XlPhoneticAlignmentToString = "xlPhoneticAlignNoControl"
        Case xlPhoneticAlignLeft: XlPhoneticAlignmentToString = "xlPhoneticAlignLeft"
        Case xlPhoneticAlignCenter: XlPhoneticAlignmentToString = "xlPhoneticAlignCenter"
        Case xlPhoneticAlignDistributed: XlPhoneticAlignmentToString = "xlPhoneticAlignDistributed"
    End Select
End Function
