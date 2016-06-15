Attribute VB_Name = "wMsoDocProperties"
Function MsoDocPropertiesFromString(value As String) As MsoDocProperties
    If IsNumeric(value) Then
        MsoDocPropertiesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoPropertyTypeNumber": MsoDocPropertiesFromString = msoPropertyTypeNumber
        Case "msoPropertyTypeBoolean": MsoDocPropertiesFromString = msoPropertyTypeBoolean
        Case "msoPropertyTypeDate": MsoDocPropertiesFromString = msoPropertyTypeDate
        Case "msoPropertyTypeString": MsoDocPropertiesFromString = msoPropertyTypeString
        Case "msoPropertyTypeFloat": MsoDocPropertiesFromString = msoPropertyTypeFloat
    End Select
End Function

Function MsoDocPropertiesToString(value As MsoDocProperties) As String
    Select Case value
        Case msoPropertyTypeNumber: MsoDocPropertiesToString = "msoPropertyTypeNumber"
        Case msoPropertyTypeBoolean: MsoDocPropertiesToString = "msoPropertyTypeBoolean"
        Case msoPropertyTypeDate: MsoDocPropertiesToString = "msoPropertyTypeDate"
        Case msoPropertyTypeString: MsoDocPropertiesToString = "msoPropertyTypeString"
        Case msoPropertyTypeFloat: MsoDocPropertiesToString = "msoPropertyTypeFloat"
    End Select
End Function
