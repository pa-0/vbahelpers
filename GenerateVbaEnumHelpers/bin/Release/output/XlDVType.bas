Attribute VB_Name = "wXlDVType"
Function XlDVTypeFromString(value As String) As XlDVType
    If IsNumeric(value) Then
        XlDVTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlValidateInputOnly": XlDVTypeFromString = xlValidateInputOnly
        Case "xlValidateWholeNumber": XlDVTypeFromString = xlValidateWholeNumber
        Case "xlValidateDecimal": XlDVTypeFromString = xlValidateDecimal
        Case "xlValidateList": XlDVTypeFromString = xlValidateList
        Case "xlValidateDate": XlDVTypeFromString = xlValidateDate
        Case "xlValidateTime": XlDVTypeFromString = xlValidateTime
        Case "xlValidateTextLength": XlDVTypeFromString = xlValidateTextLength
        Case "xlValidateCustom": XlDVTypeFromString = xlValidateCustom
    End Select
End Function

Function XlDVTypeToString(value As XlDVType) As String
    Select Case value
        Case xlValidateInputOnly: XlDVTypeToString = "xlValidateInputOnly"
        Case xlValidateWholeNumber: XlDVTypeToString = "xlValidateWholeNumber"
        Case xlValidateDecimal: XlDVTypeToString = "xlValidateDecimal"
        Case xlValidateList: XlDVTypeToString = "xlValidateList"
        Case xlValidateDate: XlDVTypeToString = "xlValidateDate"
        Case xlValidateTime: XlDVTypeToString = "xlValidateTime"
        Case xlValidateTextLength: XlDVTypeToString = "xlValidateTextLength"
        Case xlValidateCustom: XlDVTypeToString = "xlValidateCustom"
    End Select
End Function
