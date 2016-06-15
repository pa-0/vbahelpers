Attribute VB_Name = "wOlContactPhoneNumber"
Function OlContactPhoneNumberFromString(value As String) As OlContactPhoneNumber
    If IsNumeric(value) Then
        OlContactPhoneNumberFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olContactPhoneAssistant": OlContactPhoneNumberFromString = olContactPhoneAssistant
        Case "olContactPhoneBusiness": OlContactPhoneNumberFromString = olContactPhoneBusiness
        Case "olContactPhoneBusiness2": OlContactPhoneNumberFromString = olContactPhoneBusiness2
        Case "olContactPhoneBusinessFax": OlContactPhoneNumberFromString = olContactPhoneBusinessFax
        Case "olContactPhoneCallback": OlContactPhoneNumberFromString = olContactPhoneCallback
        Case "olContactPhoneCar": OlContactPhoneNumberFromString = olContactPhoneCar
        Case "olContactPhoneCompany": OlContactPhoneNumberFromString = olContactPhoneCompany
        Case "olContactPhoneHome": OlContactPhoneNumberFromString = olContactPhoneHome
        Case "olContactPhoneHome2": OlContactPhoneNumberFromString = olContactPhoneHome2
        Case "olContactPhoneHomeFax": OlContactPhoneNumberFromString = olContactPhoneHomeFax
        Case "olContactPhoneISDN": OlContactPhoneNumberFromString = olContactPhoneISDN
        Case "olContactPhoneMobile": OlContactPhoneNumberFromString = olContactPhoneMobile
        Case "olContactPhoneOther": OlContactPhoneNumberFromString = olContactPhoneOther
        Case "olContactPhoneOtherFax": OlContactPhoneNumberFromString = olContactPhoneOtherFax
        Case "olContactPhonePager": OlContactPhoneNumberFromString = olContactPhonePager
        Case "olContactPhonePrimary": OlContactPhoneNumberFromString = olContactPhonePrimary
        Case "olContactPhoneRadio": OlContactPhoneNumberFromString = olContactPhoneRadio
        Case "olContactPhoneTelex": OlContactPhoneNumberFromString = olContactPhoneTelex
        Case "olContactPhoneTTYTTD": OlContactPhoneNumberFromString = olContactPhoneTTYTTD
    End Select
End Function

Function OlContactPhoneNumberToString(value As OlContactPhoneNumber) As String
    Select Case value
        Case olContactPhoneAssistant: OlContactPhoneNumberToString = "olContactPhoneAssistant"
        Case olContactPhoneBusiness: OlContactPhoneNumberToString = "olContactPhoneBusiness"
        Case olContactPhoneBusiness2: OlContactPhoneNumberToString = "olContactPhoneBusiness2"
        Case olContactPhoneBusinessFax: OlContactPhoneNumberToString = "olContactPhoneBusinessFax"
        Case olContactPhoneCallback: OlContactPhoneNumberToString = "olContactPhoneCallback"
        Case olContactPhoneCar: OlContactPhoneNumberToString = "olContactPhoneCar"
        Case olContactPhoneCompany: OlContactPhoneNumberToString = "olContactPhoneCompany"
        Case olContactPhoneHome: OlContactPhoneNumberToString = "olContactPhoneHome"
        Case olContactPhoneHome2: OlContactPhoneNumberToString = "olContactPhoneHome2"
        Case olContactPhoneHomeFax: OlContactPhoneNumberToString = "olContactPhoneHomeFax"
        Case olContactPhoneISDN: OlContactPhoneNumberToString = "olContactPhoneISDN"
        Case olContactPhoneMobile: OlContactPhoneNumberToString = "olContactPhoneMobile"
        Case olContactPhoneOther: OlContactPhoneNumberToString = "olContactPhoneOther"
        Case olContactPhoneOtherFax: OlContactPhoneNumberToString = "olContactPhoneOtherFax"
        Case olContactPhonePager: OlContactPhoneNumberToString = "olContactPhonePager"
        Case olContactPhonePrimary: OlContactPhoneNumberToString = "olContactPhonePrimary"
        Case olContactPhoneRadio: OlContactPhoneNumberToString = "olContactPhoneRadio"
        Case olContactPhoneTelex: OlContactPhoneNumberToString = "olContactPhoneTelex"
        Case olContactPhoneTTYTTD: OlContactPhoneNumberToString = "olContactPhoneTTYTTD"
    End Select
End Function
