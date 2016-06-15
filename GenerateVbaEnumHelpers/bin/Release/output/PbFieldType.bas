Attribute VB_Name = "wPbFieldType"
Function PbFieldTypeFromString(value As String) As PbFieldType
    If IsNumeric(value) Then
        PbFieldTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFieldNone": PbFieldTypeFromString = pbFieldNone
        Case "pbFieldPageNumber": PbFieldTypeFromString = pbFieldPageNumber
        Case "pbFieldPageNumberNext": PbFieldTypeFromString = pbFieldPageNumberNext
        Case "pbFieldPageNumberPrev": PbFieldTypeFromString = pbFieldPageNumberPrev
        Case "pbFieldDateTime": PbFieldTypeFromString = pbFieldDateTime
        Case "pbFieldMailMerge": PbFieldTypeFromString = pbFieldMailMerge
        Case "pbFieldIHIV": PbFieldTypeFromString = pbFieldIHIV
        Case "pbFieldPhoneticGuide": PbFieldTypeFromString = pbFieldPhoneticGuide
        Case "pbFieldWizardSampleText": PbFieldTypeFromString = pbFieldWizardSampleText
        Case "pbFieldHyperlinkURL": PbFieldTypeFromString = pbFieldHyperlinkURL
        Case "pbFieldHyperlinkRelativePage": PbFieldTypeFromString = pbFieldHyperlinkRelativePage
        Case "pbFieldHyperlinkAbsolutePage": PbFieldTypeFromString = pbFieldHyperlinkAbsolutePage
        Case "pbFieldHyperlinkEmail": PbFieldTypeFromString = pbFieldHyperlinkEmail
        Case "pbFieldHyperlinkFile": PbFieldTypeFromString = pbFieldHyperlinkFile
        Case "pbFieldPersonalizedHyperlinkURL": PbFieldTypeFromString = pbFieldPersonalizedHyperlinkURL
    End Select
End Function

Function PbFieldTypeToString(value As PbFieldType) As String
    Select Case value
        Case pbFieldNone: PbFieldTypeToString = "pbFieldNone"
        Case pbFieldPageNumber: PbFieldTypeToString = "pbFieldPageNumber"
        Case pbFieldPageNumberNext: PbFieldTypeToString = "pbFieldPageNumberNext"
        Case pbFieldPageNumberPrev: PbFieldTypeToString = "pbFieldPageNumberPrev"
        Case pbFieldDateTime: PbFieldTypeToString = "pbFieldDateTime"
        Case pbFieldMailMerge: PbFieldTypeToString = "pbFieldMailMerge"
        Case pbFieldIHIV: PbFieldTypeToString = "pbFieldIHIV"
        Case pbFieldPhoneticGuide: PbFieldTypeToString = "pbFieldPhoneticGuide"
        Case pbFieldWizardSampleText: PbFieldTypeToString = "pbFieldWizardSampleText"
        Case pbFieldHyperlinkURL: PbFieldTypeToString = "pbFieldHyperlinkURL"
        Case pbFieldHyperlinkRelativePage: PbFieldTypeToString = "pbFieldHyperlinkRelativePage"
        Case pbFieldHyperlinkAbsolutePage: PbFieldTypeToString = "pbFieldHyperlinkAbsolutePage"
        Case pbFieldHyperlinkEmail: PbFieldTypeToString = "pbFieldHyperlinkEmail"
        Case pbFieldHyperlinkFile: PbFieldTypeToString = "pbFieldHyperlinkFile"
        Case pbFieldPersonalizedHyperlinkURL: PbFieldTypeToString = "pbFieldPersonalizedHyperlinkURL"
    End Select
End Function
