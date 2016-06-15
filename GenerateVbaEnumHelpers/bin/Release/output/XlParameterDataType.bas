Attribute VB_Name = "wXlParameterDataType"
Function XlParameterDataTypeFromString(value As String) As XlParameterDataType
    If IsNumeric(value) Then
        XlParameterDataTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlParamTypeUnknown": XlParameterDataTypeFromString = xlParamTypeUnknown
        Case "xlParamTypeChar": XlParameterDataTypeFromString = xlParamTypeChar
        Case "xlParamTypeNumeric": XlParameterDataTypeFromString = xlParamTypeNumeric
        Case "xlParamTypeDecimal": XlParameterDataTypeFromString = xlParamTypeDecimal
        Case "xlParamTypeInteger": XlParameterDataTypeFromString = xlParamTypeInteger
        Case "xlParamTypeSmallInt": XlParameterDataTypeFromString = xlParamTypeSmallInt
        Case "xlParamTypeFloat": XlParameterDataTypeFromString = xlParamTypeFloat
        Case "xlParamTypeReal": XlParameterDataTypeFromString = xlParamTypeReal
        Case "xlParamTypeDouble": XlParameterDataTypeFromString = xlParamTypeDouble
        Case "xlParamTypeDate": XlParameterDataTypeFromString = xlParamTypeDate
        Case "xlParamTypeTime": XlParameterDataTypeFromString = xlParamTypeTime
        Case "xlParamTypeTimestamp": XlParameterDataTypeFromString = xlParamTypeTimestamp
        Case "xlParamTypeVarChar": XlParameterDataTypeFromString = xlParamTypeVarChar
        Case "xlParamTypeWChar": XlParameterDataTypeFromString = xlParamTypeWChar
        Case "xlParamTypeBit": XlParameterDataTypeFromString = xlParamTypeBit
        Case "xlParamTypeTinyInt": XlParameterDataTypeFromString = xlParamTypeTinyInt
        Case "xlParamTypeBigInt": XlParameterDataTypeFromString = xlParamTypeBigInt
        Case "xlParamTypeLongVarBinary": XlParameterDataTypeFromString = xlParamTypeLongVarBinary
        Case "xlParamTypeVarBinary": XlParameterDataTypeFromString = xlParamTypeVarBinary
        Case "xlParamTypeBinary": XlParameterDataTypeFromString = xlParamTypeBinary
        Case "xlParamTypeLongVarChar": XlParameterDataTypeFromString = xlParamTypeLongVarChar
    End Select
End Function

Function XlParameterDataTypeToString(value As XlParameterDataType) As String
    Select Case value
        Case xlParamTypeUnknown: XlParameterDataTypeToString = "xlParamTypeUnknown"
        Case xlParamTypeChar: XlParameterDataTypeToString = "xlParamTypeChar"
        Case xlParamTypeNumeric: XlParameterDataTypeToString = "xlParamTypeNumeric"
        Case xlParamTypeDecimal: XlParameterDataTypeToString = "xlParamTypeDecimal"
        Case xlParamTypeInteger: XlParameterDataTypeToString = "xlParamTypeInteger"
        Case xlParamTypeSmallInt: XlParameterDataTypeToString = "xlParamTypeSmallInt"
        Case xlParamTypeFloat: XlParameterDataTypeToString = "xlParamTypeFloat"
        Case xlParamTypeReal: XlParameterDataTypeToString = "xlParamTypeReal"
        Case xlParamTypeDouble: XlParameterDataTypeToString = "xlParamTypeDouble"
        Case xlParamTypeDate: XlParameterDataTypeToString = "xlParamTypeDate"
        Case xlParamTypeTime: XlParameterDataTypeToString = "xlParamTypeTime"
        Case xlParamTypeTimestamp: XlParameterDataTypeToString = "xlParamTypeTimestamp"
        Case xlParamTypeVarChar: XlParameterDataTypeToString = "xlParamTypeVarChar"
        Case xlParamTypeWChar: XlParameterDataTypeToString = "xlParamTypeWChar"
        Case xlParamTypeBit: XlParameterDataTypeToString = "xlParamTypeBit"
        Case xlParamTypeTinyInt: XlParameterDataTypeToString = "xlParamTypeTinyInt"
        Case xlParamTypeBigInt: XlParameterDataTypeToString = "xlParamTypeBigInt"
        Case xlParamTypeLongVarBinary: XlParameterDataTypeToString = "xlParamTypeLongVarBinary"
        Case xlParamTypeVarBinary: XlParameterDataTypeToString = "xlParamTypeVarBinary"
        Case xlParamTypeBinary: XlParameterDataTypeToString = "xlParamTypeBinary"
        Case xlParamTypeLongVarChar: XlParameterDataTypeToString = "xlParamTypeLongVarChar"
    End Select
End Function
