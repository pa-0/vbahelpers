Attribute VB_Name = "wPbPersonalInfoSet"
Function PbPersonalInfoSetFromString(value As String) As PbPersonalInfoSet
    If IsNumeric(value) Then
        PbPersonalInfoSetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPersonalInfoPrimaryBusiness": PbPersonalInfoSetFromString = pbPersonalInfoPrimaryBusiness
        Case "pbPersonalInfoSecondaryBusiness": PbPersonalInfoSetFromString = pbPersonalInfoSecondaryBusiness
        Case "pbPersonalInfoOtherOrganization": PbPersonalInfoSetFromString = pbPersonalInfoOtherOrganization
        Case "pbPersonalInfoHome": PbPersonalInfoSetFromString = pbPersonalInfoHome
    End Select
End Function

Function PbPersonalInfoSetToString(value As PbPersonalInfoSet) As String
    Select Case value
        Case pbPersonalInfoPrimaryBusiness: PbPersonalInfoSetToString = "pbPersonalInfoPrimaryBusiness"
        Case pbPersonalInfoSecondaryBusiness: PbPersonalInfoSetToString = "pbPersonalInfoSecondaryBusiness"
        Case pbPersonalInfoOtherOrganization: PbPersonalInfoSetToString = "pbPersonalInfoOtherOrganization"
        Case pbPersonalInfoHome: PbPersonalInfoSetToString = "pbPersonalInfoHome"
    End Select
End Function
