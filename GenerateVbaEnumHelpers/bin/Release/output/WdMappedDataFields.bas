Attribute VB_Name = "wWdMappedDataFields"
Function WdMappedDataFieldsFromString(value As String) As WdMappedDataFields
    If IsNumeric(value) Then
        WdMappedDataFieldsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdUniqueIdentifier": WdMappedDataFieldsFromString = wdUniqueIdentifier
        Case "wdCourtesyTitle": WdMappedDataFieldsFromString = wdCourtesyTitle
        Case "wdFirstName": WdMappedDataFieldsFromString = wdFirstName
        Case "wdMiddleName": WdMappedDataFieldsFromString = wdMiddleName
        Case "wdLastName": WdMappedDataFieldsFromString = wdLastName
        Case "wdSuffix": WdMappedDataFieldsFromString = wdSuffix
        Case "wdNickname": WdMappedDataFieldsFromString = wdNickname
        Case "wdJobTitle": WdMappedDataFieldsFromString = wdJobTitle
        Case "wdCompany": WdMappedDataFieldsFromString = wdCompany
        Case "wdAddress1": WdMappedDataFieldsFromString = wdAddress1
        Case "wdAddress2": WdMappedDataFieldsFromString = wdAddress2
        Case "wdCity": WdMappedDataFieldsFromString = wdCity
        Case "wdState": WdMappedDataFieldsFromString = wdState
        Case "wdPostalCode": WdMappedDataFieldsFromString = wdPostalCode
        Case "wdCountryRegion": WdMappedDataFieldsFromString = wdCountryRegion
        Case "wdBusinessPhone": WdMappedDataFieldsFromString = wdBusinessPhone
        Case "wdBusinessFax": WdMappedDataFieldsFromString = wdBusinessFax
        Case "wdHomePhone": WdMappedDataFieldsFromString = wdHomePhone
        Case "wdHomeFax": WdMappedDataFieldsFromString = wdHomeFax
        Case "wdEmailAddress": WdMappedDataFieldsFromString = wdEmailAddress
        Case "wdWebPageURL": WdMappedDataFieldsFromString = wdWebPageURL
        Case "wdSpouseCourtesyTitle": WdMappedDataFieldsFromString = wdSpouseCourtesyTitle
        Case "wdSpouseFirstName": WdMappedDataFieldsFromString = wdSpouseFirstName
        Case "wdSpouseMiddleName": WdMappedDataFieldsFromString = wdSpouseMiddleName
        Case "wdSpouseLastName": WdMappedDataFieldsFromString = wdSpouseLastName
        Case "wdSpouseNickname": WdMappedDataFieldsFromString = wdSpouseNickname
        Case "wdRubyFirstName": WdMappedDataFieldsFromString = wdRubyFirstName
        Case "wdRubyLastName": WdMappedDataFieldsFromString = wdRubyLastName
        Case "wdAddress3": WdMappedDataFieldsFromString = wdAddress3
        Case "wdDepartment": WdMappedDataFieldsFromString = wdDepartment
    End Select
End Function

Function WdMappedDataFieldsToString(value As WdMappedDataFields) As String
    Select Case value
        Case wdUniqueIdentifier: WdMappedDataFieldsToString = "wdUniqueIdentifier"
        Case wdCourtesyTitle: WdMappedDataFieldsToString = "wdCourtesyTitle"
        Case wdFirstName: WdMappedDataFieldsToString = "wdFirstName"
        Case wdMiddleName: WdMappedDataFieldsToString = "wdMiddleName"
        Case wdLastName: WdMappedDataFieldsToString = "wdLastName"
        Case wdSuffix: WdMappedDataFieldsToString = "wdSuffix"
        Case wdNickname: WdMappedDataFieldsToString = "wdNickname"
        Case wdJobTitle: WdMappedDataFieldsToString = "wdJobTitle"
        Case wdCompany: WdMappedDataFieldsToString = "wdCompany"
        Case wdAddress1: WdMappedDataFieldsToString = "wdAddress1"
        Case wdAddress2: WdMappedDataFieldsToString = "wdAddress2"
        Case wdCity: WdMappedDataFieldsToString = "wdCity"
        Case wdState: WdMappedDataFieldsToString = "wdState"
        Case wdPostalCode: WdMappedDataFieldsToString = "wdPostalCode"
        Case wdCountryRegion: WdMappedDataFieldsToString = "wdCountryRegion"
        Case wdBusinessPhone: WdMappedDataFieldsToString = "wdBusinessPhone"
        Case wdBusinessFax: WdMappedDataFieldsToString = "wdBusinessFax"
        Case wdHomePhone: WdMappedDataFieldsToString = "wdHomePhone"
        Case wdHomeFax: WdMappedDataFieldsToString = "wdHomeFax"
        Case wdEmailAddress: WdMappedDataFieldsToString = "wdEmailAddress"
        Case wdWebPageURL: WdMappedDataFieldsToString = "wdWebPageURL"
        Case wdSpouseCourtesyTitle: WdMappedDataFieldsToString = "wdSpouseCourtesyTitle"
        Case wdSpouseFirstName: WdMappedDataFieldsToString = "wdSpouseFirstName"
        Case wdSpouseMiddleName: WdMappedDataFieldsToString = "wdSpouseMiddleName"
        Case wdSpouseLastName: WdMappedDataFieldsToString = "wdSpouseLastName"
        Case wdSpouseNickname: WdMappedDataFieldsToString = "wdSpouseNickname"
        Case wdRubyFirstName: WdMappedDataFieldsToString = "wdRubyFirstName"
        Case wdRubyLastName: WdMappedDataFieldsToString = "wdRubyLastName"
        Case wdAddress3: WdMappedDataFieldsToString = "wdAddress3"
        Case wdDepartment: WdMappedDataFieldsToString = "wdDepartment"
    End Select
End Function
