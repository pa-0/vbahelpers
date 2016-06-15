Attribute VB_Name = "wPbMappedDataFields"
Function PbMappedDataFieldsFromString(value As String) As PbMappedDataFields
    If IsNumeric(value) Then
        PbMappedDataFieldsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbUniqueIdentifier": PbMappedDataFieldsFromString = pbUniqueIdentifier
        Case "pbCourtesyTitle": PbMappedDataFieldsFromString = pbCourtesyTitle
        Case "pbFirstName": PbMappedDataFieldsFromString = pbFirstName
        Case "pbMiddleName": PbMappedDataFieldsFromString = pbMiddleName
        Case "pbLastName": PbMappedDataFieldsFromString = pbLastName
        Case "pbSuffix": PbMappedDataFieldsFromString = pbSuffix
        Case "pbNickname": PbMappedDataFieldsFromString = pbNickname
        Case "pbJobTitle": PbMappedDataFieldsFromString = pbJobTitle
        Case "pbCompany": PbMappedDataFieldsFromString = pbCompany
        Case "pbAddress1": PbMappedDataFieldsFromString = pbAddress1
        Case "pbAddress2": PbMappedDataFieldsFromString = pbAddress2
        Case "pbCity": PbMappedDataFieldsFromString = pbCity
        Case "pbState": PbMappedDataFieldsFromString = pbState
        Case "pbPostalCode": PbMappedDataFieldsFromString = pbPostalCode
        Case "pbCountryRegion": PbMappedDataFieldsFromString = pbCountryRegion
        Case "pbBusinessPhone": PbMappedDataFieldsFromString = pbBusinessPhone
        Case "pbBusinessFax": PbMappedDataFieldsFromString = pbBusinessFax
        Case "pbHomePhone": PbMappedDataFieldsFromString = pbHomePhone
        Case "pbHomeFax": PbMappedDataFieldsFromString = pbHomeFax
        Case "pbEmailAddress": PbMappedDataFieldsFromString = pbEmailAddress
        Case "pbWebPageURL": PbMappedDataFieldsFromString = pbWebPageURL
        Case "pbSpouseCourtesyTitle": PbMappedDataFieldsFromString = pbSpouseCourtesyTitle
        Case "pbSpouseFirstName": PbMappedDataFieldsFromString = pbSpouseFirstName
        Case "pbSpouseMiddleName": PbMappedDataFieldsFromString = pbSpouseMiddleName
        Case "pbSpouseLastName": PbMappedDataFieldsFromString = pbSpouseLastName
        Case "pbSpouseNickname": PbMappedDataFieldsFromString = pbSpouseNickname
        Case "pbRubyFirstName": PbMappedDataFieldsFromString = pbRubyFirstName
        Case "pbRubyLastName": PbMappedDataFieldsFromString = pbRubyLastName
        Case "pbAddress3": PbMappedDataFieldsFromString = pbAddress3
        Case "pbDepartment": PbMappedDataFieldsFromString = pbDepartment
    End Select
End Function

Function PbMappedDataFieldsToString(value As PbMappedDataFields) As String
    Select Case value
        Case pbUniqueIdentifier: PbMappedDataFieldsToString = "pbUniqueIdentifier"
        Case pbCourtesyTitle: PbMappedDataFieldsToString = "pbCourtesyTitle"
        Case pbFirstName: PbMappedDataFieldsToString = "pbFirstName"
        Case pbMiddleName: PbMappedDataFieldsToString = "pbMiddleName"
        Case pbLastName: PbMappedDataFieldsToString = "pbLastName"
        Case pbSuffix: PbMappedDataFieldsToString = "pbSuffix"
        Case pbNickname: PbMappedDataFieldsToString = "pbNickname"
        Case pbJobTitle: PbMappedDataFieldsToString = "pbJobTitle"
        Case pbCompany: PbMappedDataFieldsToString = "pbCompany"
        Case pbAddress1: PbMappedDataFieldsToString = "pbAddress1"
        Case pbAddress2: PbMappedDataFieldsToString = "pbAddress2"
        Case pbCity: PbMappedDataFieldsToString = "pbCity"
        Case pbState: PbMappedDataFieldsToString = "pbState"
        Case pbPostalCode: PbMappedDataFieldsToString = "pbPostalCode"
        Case pbCountryRegion: PbMappedDataFieldsToString = "pbCountryRegion"
        Case pbBusinessPhone: PbMappedDataFieldsToString = "pbBusinessPhone"
        Case pbBusinessFax: PbMappedDataFieldsToString = "pbBusinessFax"
        Case pbHomePhone: PbMappedDataFieldsToString = "pbHomePhone"
        Case pbHomeFax: PbMappedDataFieldsToString = "pbHomeFax"
        Case pbEmailAddress: PbMappedDataFieldsToString = "pbEmailAddress"
        Case pbWebPageURL: PbMappedDataFieldsToString = "pbWebPageURL"
        Case pbSpouseCourtesyTitle: PbMappedDataFieldsToString = "pbSpouseCourtesyTitle"
        Case pbSpouseFirstName: PbMappedDataFieldsToString = "pbSpouseFirstName"
        Case pbSpouseMiddleName: PbMappedDataFieldsToString = "pbSpouseMiddleName"
        Case pbSpouseLastName: PbMappedDataFieldsToString = "pbSpouseLastName"
        Case pbSpouseNickname: PbMappedDataFieldsToString = "pbSpouseNickname"
        Case pbRubyFirstName: PbMappedDataFieldsToString = "pbRubyFirstName"
        Case pbRubyLastName: PbMappedDataFieldsToString = "pbRubyLastName"
        Case pbAddress3: PbMappedDataFieldsToString = "pbAddress3"
        Case pbDepartment: PbMappedDataFieldsToString = "pbDepartment"
    End Select
End Function
