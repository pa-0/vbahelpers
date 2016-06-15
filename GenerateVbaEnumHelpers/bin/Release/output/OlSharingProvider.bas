Attribute VB_Name = "wOlSharingProvider"
Function OlSharingProviderFromString(value As String) As OlSharingProvider
    If IsNumeric(value) Then
        OlSharingProviderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olProviderUnknown": OlSharingProviderFromString = olProviderUnknown
        Case "olProviderExchange": OlSharingProviderFromString = olProviderExchange
        Case "olProviderWebCal": OlSharingProviderFromString = olProviderWebCal
        Case "olProviderPubCal": OlSharingProviderFromString = olProviderPubCal
        Case "olProviderICal": OlSharingProviderFromString = olProviderICal
        Case "olProviderSharePoint": OlSharingProviderFromString = olProviderSharePoint
        Case "olProviderRSS": OlSharingProviderFromString = olProviderRSS
        Case "olProviderFederate": OlSharingProviderFromString = olProviderFederate
    End Select
End Function

Function OlSharingProviderToString(value As OlSharingProvider) As String
    Select Case value
        Case olProviderUnknown: OlSharingProviderToString = "olProviderUnknown"
        Case olProviderExchange: OlSharingProviderToString = "olProviderExchange"
        Case olProviderWebCal: OlSharingProviderToString = "olProviderWebCal"
        Case olProviderPubCal: OlSharingProviderToString = "olProviderPubCal"
        Case olProviderICal: OlSharingProviderToString = "olProviderICal"
        Case olProviderSharePoint: OlSharingProviderToString = "olProviderSharePoint"
        Case olProviderRSS: OlSharingProviderToString = "olProviderRSS"
        Case olProviderFederate: OlSharingProviderToString = "olProviderFederate"
    End Select
End Function
