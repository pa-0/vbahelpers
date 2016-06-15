Attribute VB_Name = "wMsoContactCardAddressType"
Function MsoContactCardAddressTypeFromString(value As String) As MsoContactCardAddressType
    If IsNumeric(value) Then
        MsoContactCardAddressTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoContactCardAddressTypeUnknown": MsoContactCardAddressTypeFromString = msoContactCardAddressTypeUnknown
        Case "msoContactCardAddressTypeOutlook": MsoContactCardAddressTypeFromString = msoContactCardAddressTypeOutlook
        Case "msoContactCardAddressTypeSMTP": MsoContactCardAddressTypeFromString = msoContactCardAddressTypeSMTP
        Case "msoContactCardAddressTypeIM": MsoContactCardAddressTypeFromString = msoContactCardAddressTypeIM
    End Select
End Function

Function MsoContactCardAddressTypeToString(value As MsoContactCardAddressType) As String
    Select Case value
        Case msoContactCardAddressTypeUnknown: MsoContactCardAddressTypeToString = "msoContactCardAddressTypeUnknown"
        Case msoContactCardAddressTypeOutlook: MsoContactCardAddressTypeToString = "msoContactCardAddressTypeOutlook"
        Case msoContactCardAddressTypeSMTP: MsoContactCardAddressTypeToString = "msoContactCardAddressTypeSMTP"
        Case msoContactCardAddressTypeIM: MsoContactCardAddressTypeToString = "msoContactCardAddressTypeIM"
    End Select
End Function
