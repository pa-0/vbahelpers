Attribute VB_Name = "wXlDVAlertStyle"
Function XlDVAlertStyleFromString(value As String) As XlDVAlertStyle
    If IsNumeric(value) Then
        XlDVAlertStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlValidAlertStop": XlDVAlertStyleFromString = xlValidAlertStop
        Case "xlValidAlertWarning": XlDVAlertStyleFromString = xlValidAlertWarning
        Case "xlValidAlertInformation": XlDVAlertStyleFromString = xlValidAlertInformation
    End Select
End Function

Function XlDVAlertStyleToString(value As XlDVAlertStyle) As String
    Select Case value
        Case xlValidAlertStop: XlDVAlertStyleToString = "xlValidAlertStop"
        Case xlValidAlertWarning: XlDVAlertStyleToString = "xlValidAlertWarning"
        Case xlValidAlertInformation: XlDVAlertStyleToString = "xlValidAlertInformation"
    End Select
End Function
