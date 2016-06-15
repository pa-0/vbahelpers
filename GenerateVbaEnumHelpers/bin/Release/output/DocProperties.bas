Attribute VB_Name = "wDocProperties"
Function DocPropertiesFromString(value As String) As DocProperties
    If IsNumeric(value) Then
        DocPropertiesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "offPropertyTypeNumber": DocPropertiesFromString = offPropertyTypeNumber
        Case "offPropertyTypeBoolean": DocPropertiesFromString = offPropertyTypeBoolean
        Case "offPropertyTypeDate": DocPropertiesFromString = offPropertyTypeDate
        Case "offPropertyTypeString": DocPropertiesFromString = offPropertyTypeString
        Case "offPropertyTypeFloat": DocPropertiesFromString = offPropertyTypeFloat
    End Select
End Function

Function DocPropertiesToString(value As DocProperties) As String
    Select Case value
        Case offPropertyTypeNumber: DocPropertiesToString = "offPropertyTypeNumber"
        Case offPropertyTypeBoolean: DocPropertiesToString = "offPropertyTypeBoolean"
        Case offPropertyTypeDate: DocPropertiesToString = "offPropertyTypeDate"
        Case offPropertyTypeString: DocPropertiesToString = "offPropertyTypeString"
        Case offPropertyTypeFloat: DocPropertiesToString = "offPropertyTypeFloat"
    End Select
End Function
