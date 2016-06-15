Attribute VB_Name = "wMsoPickerField"
Function MsoPickerFieldFromString(value As String) As MsoPickerField
    If IsNumeric(value) Then
        MsoPickerFieldFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoPickerFieldUnknown": MsoPickerFieldFromString = msoPickerFieldUnknown
        Case "msoPickerFieldDateTime": MsoPickerFieldFromString = msoPickerFieldDateTime
        Case "msoPickerFieldNumber": MsoPickerFieldFromString = msoPickerFieldNumber
        Case "msoPickerFieldText": MsoPickerFieldFromString = msoPickerFieldText
        Case "msoPickerFieldUser": MsoPickerFieldFromString = msoPickerFieldUser
        Case "msoPickerFieldMax": MsoPickerFieldFromString = msoPickerFieldMax
    End Select
End Function

Function MsoPickerFieldToString(value As MsoPickerField) As String
    Select Case value
        Case msoPickerFieldUnknown: MsoPickerFieldToString = "msoPickerFieldUnknown"
        Case msoPickerFieldDateTime: MsoPickerFieldToString = "msoPickerFieldDateTime"
        Case msoPickerFieldNumber: MsoPickerFieldToString = "msoPickerFieldNumber"
        Case msoPickerFieldText: MsoPickerFieldToString = "msoPickerFieldText"
        Case msoPickerFieldUser: MsoPickerFieldToString = "msoPickerFieldUser"
        Case msoPickerFieldMax: MsoPickerFieldToString = "msoPickerFieldMax"
    End Select
End Function
