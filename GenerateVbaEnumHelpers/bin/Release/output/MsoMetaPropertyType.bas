Attribute VB_Name = "wMsoMetaPropertyType"
Function MsoMetaPropertyTypeFromString(value As String) As MsoMetaPropertyType
    If IsNumeric(value) Then
        MsoMetaPropertyTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoMetaPropertyTypeUnknown": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeUnknown
        Case "msoMetaPropertyTypeBoolean": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeBoolean
        Case "msoMetaPropertyTypeChoice": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeChoice
        Case "msoMetaPropertyTypeCalculated": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeCalculated
        Case "msoMetaPropertyTypeComputed": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeComputed
        Case "msoMetaPropertyTypeCurrency": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeCurrency
        Case "msoMetaPropertyTypeDateTime": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeDateTime
        Case "msoMetaPropertyTypeFillInChoice": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeFillInChoice
        Case "msoMetaPropertyTypeGuid": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeGuid
        Case "msoMetaPropertyTypeInteger": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeInteger
        Case "msoMetaPropertyTypeLookup": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeLookup
        Case "msoMetaPropertyTypeMultiChoiceLookup": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeMultiChoiceLookup
        Case "msoMetaPropertyTypeMultiChoice": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeMultiChoice
        Case "msoMetaPropertyTypeMultiChoiceFillIn": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeMultiChoiceFillIn
        Case "msoMetaPropertyTypeNote": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeNote
        Case "msoMetaPropertyTypeNumber": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeNumber
        Case "msoMetaPropertyTypeText": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeText
        Case "msoMetaPropertyTypeUrl": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeUrl
        Case "msoMetaPropertyTypeUser": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeUser
        Case "msoMetaPropertyTypeUserMulti": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeUserMulti
        Case "msoMetaPropertyTypeBusinessData": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeBusinessData
        Case "msoMetaPropertyTypeBusinessDataSecondary": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeBusinessDataSecondary
        Case "msoMetaPropertyTypeMax": MsoMetaPropertyTypeFromString = msoMetaPropertyTypeMax
    End Select
End Function

Function MsoMetaPropertyTypeToString(value As MsoMetaPropertyType) As String
    Select Case value
        Case msoMetaPropertyTypeUnknown: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeUnknown"
        Case msoMetaPropertyTypeBoolean: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeBoolean"
        Case msoMetaPropertyTypeChoice: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeChoice"
        Case msoMetaPropertyTypeCalculated: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeCalculated"
        Case msoMetaPropertyTypeComputed: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeComputed"
        Case msoMetaPropertyTypeCurrency: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeCurrency"
        Case msoMetaPropertyTypeDateTime: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeDateTime"
        Case msoMetaPropertyTypeFillInChoice: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeFillInChoice"
        Case msoMetaPropertyTypeGuid: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeGuid"
        Case msoMetaPropertyTypeInteger: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeInteger"
        Case msoMetaPropertyTypeLookup: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeLookup"
        Case msoMetaPropertyTypeMultiChoiceLookup: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeMultiChoiceLookup"
        Case msoMetaPropertyTypeMultiChoice: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeMultiChoice"
        Case msoMetaPropertyTypeMultiChoiceFillIn: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeMultiChoiceFillIn"
        Case msoMetaPropertyTypeNote: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeNote"
        Case msoMetaPropertyTypeNumber: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeNumber"
        Case msoMetaPropertyTypeText: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeText"
        Case msoMetaPropertyTypeUrl: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeUrl"
        Case msoMetaPropertyTypeUser: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeUser"
        Case msoMetaPropertyTypeUserMulti: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeUserMulti"
        Case msoMetaPropertyTypeBusinessData: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeBusinessData"
        Case msoMetaPropertyTypeBusinessDataSecondary: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeBusinessDataSecondary"
        Case msoMetaPropertyTypeMax: MsoMetaPropertyTypeToString = "msoMetaPropertyTypeMax"
    End Select
End Function
