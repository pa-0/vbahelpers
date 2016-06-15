Attribute VB_Name = "wMsoContactCardType"
Function MsoContactCardTypeFromString(value As String) As MsoContactCardType
    If IsNumeric(value) Then
        MsoContactCardTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoContactCardTypeEnterpriseContact": MsoContactCardTypeFromString = msoContactCardTypeEnterpriseContact
        Case "msoContactCardTypePersonalContact": MsoContactCardTypeFromString = msoContactCardTypePersonalContact
        Case "msoContactCardTypeUnknownContact": MsoContactCardTypeFromString = msoContactCardTypeUnknownContact
        Case "msoContactCardTypeEnterpriseGroup": MsoContactCardTypeFromString = msoContactCardTypeEnterpriseGroup
        Case "msoContactCardTypePersonalDistributionList": MsoContactCardTypeFromString = msoContactCardTypePersonalDistributionList
    End Select
End Function

Function MsoContactCardTypeToString(value As MsoContactCardType) As String
    Select Case value
        Case msoContactCardTypeEnterpriseContact: MsoContactCardTypeToString = "msoContactCardTypeEnterpriseContact"
        Case msoContactCardTypePersonalContact: MsoContactCardTypeToString = "msoContactCardTypePersonalContact"
        Case msoContactCardTypeUnknownContact: MsoContactCardTypeToString = "msoContactCardTypeUnknownContact"
        Case msoContactCardTypeEnterpriseGroup: MsoContactCardTypeToString = "msoContactCardTypeEnterpriseGroup"
        Case msoContactCardTypePersonalDistributionList: MsoContactCardTypeToString = "msoContactCardTypePersonalDistributionList"
    End Select
End Function
