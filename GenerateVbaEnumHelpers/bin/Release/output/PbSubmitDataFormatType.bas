Attribute VB_Name = "wPbSubmitDataFormatType"
Function PbSubmitDataFormatTypeFromString(value As String) As PbSubmitDataFormatType
    If IsNumeric(value) Then
        PbSubmitDataFormatTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbSubmitDataFormatHTML": PbSubmitDataFormatTypeFromString = pbSubmitDataFormatHTML
        Case "pbSubmitDataFormatRichText": PbSubmitDataFormatTypeFromString = pbSubmitDataFormatRichText
        Case "pbSubmitDataFormatCSV": PbSubmitDataFormatTypeFromString = pbSubmitDataFormatCSV
        Case "pbSubmitDataFormatTab": PbSubmitDataFormatTypeFromString = pbSubmitDataFormatTab
    End Select
End Function

Function PbSubmitDataFormatTypeToString(value As PbSubmitDataFormatType) As String
    Select Case value
        Case pbSubmitDataFormatHTML: PbSubmitDataFormatTypeToString = "pbSubmitDataFormatHTML"
        Case pbSubmitDataFormatRichText: PbSubmitDataFormatTypeToString = "pbSubmitDataFormatRichText"
        Case pbSubmitDataFormatCSV: PbSubmitDataFormatTypeToString = "pbSubmitDataFormatCSV"
        Case pbSubmitDataFormatTab: PbSubmitDataFormatTypeToString = "pbSubmitDataFormatTab"
    End Select
End Function
