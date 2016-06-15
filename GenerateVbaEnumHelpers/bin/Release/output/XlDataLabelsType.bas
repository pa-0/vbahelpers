Attribute VB_Name = "wXlDataLabelsType"
Function XlDataLabelsTypeFromString(value As String) As XlDataLabelsType
    If IsNumeric(value) Then
        XlDataLabelsTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataLabelsShowValue": XlDataLabelsTypeFromString = xlDataLabelsShowValue
        Case "xlDataLabelsShowPercent": XlDataLabelsTypeFromString = xlDataLabelsShowPercent
        Case "xlDataLabelsShowLabel": XlDataLabelsTypeFromString = xlDataLabelsShowLabel
        Case "xlDataLabelsShowLabelAndPercent": XlDataLabelsTypeFromString = xlDataLabelsShowLabelAndPercent
        Case "xlDataLabelsShowBubbleSizes": XlDataLabelsTypeFromString = xlDataLabelsShowBubbleSizes
        Case "xlDataLabelsShowNone": XlDataLabelsTypeFromString = xlDataLabelsShowNone
    End Select
End Function

Function XlDataLabelsTypeToString(value As XlDataLabelsType) As String
    Select Case value
        Case xlDataLabelsShowValue: XlDataLabelsTypeToString = "xlDataLabelsShowValue"
        Case xlDataLabelsShowPercent: XlDataLabelsTypeToString = "xlDataLabelsShowPercent"
        Case xlDataLabelsShowLabel: XlDataLabelsTypeToString = "xlDataLabelsShowLabel"
        Case xlDataLabelsShowLabelAndPercent: XlDataLabelsTypeToString = "xlDataLabelsShowLabelAndPercent"
        Case xlDataLabelsShowBubbleSizes: XlDataLabelsTypeToString = "xlDataLabelsShowBubbleSizes"
        Case xlDataLabelsShowNone: XlDataLabelsTypeToString = "xlDataLabelsShowNone"
    End Select
End Function
