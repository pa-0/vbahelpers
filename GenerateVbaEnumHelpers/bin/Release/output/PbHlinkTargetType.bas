Attribute VB_Name = "wPbHlinkTargetType"
Function PbHlinkTargetTypeFromString(value As String) As PbHlinkTargetType
    If IsNumeric(value) Then
        PbHlinkTargetTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbHlinkTargetTypeNone": PbHlinkTargetTypeFromString = pbHlinkTargetTypeNone
        Case "pbHlinkTargetTypeURL": PbHlinkTargetTypeFromString = pbHlinkTargetTypeURL
        Case "pbHlinkTargetTypeEmail": PbHlinkTargetTypeFromString = pbHlinkTargetTypeEmail
        Case "pbHlinkTargetTypeFirstPage": PbHlinkTargetTypeFromString = pbHlinkTargetTypeFirstPage
        Case "pbHlinkTargetTypeLastPage": PbHlinkTargetTypeFromString = pbHlinkTargetTypeLastPage
        Case "pbHlinkTargetTypeNextPage": PbHlinkTargetTypeFromString = pbHlinkTargetTypeNextPage
        Case "pbHlinkTargetTypePreviousPage": PbHlinkTargetTypeFromString = pbHlinkTargetTypePreviousPage
        Case "pbHlinkTargetTypePageID": PbHlinkTargetTypeFromString = pbHlinkTargetTypePageID
        Case "pbHlinkTargetTypePersonalized": PbHlinkTargetTypeFromString = pbHlinkTargetTypePersonalized
    End Select
End Function

Function PbHlinkTargetTypeToString(value As PbHlinkTargetType) As String
    Select Case value
        Case pbHlinkTargetTypeNone: PbHlinkTargetTypeToString = "pbHlinkTargetTypeNone"
        Case pbHlinkTargetTypeURL: PbHlinkTargetTypeToString = "pbHlinkTargetTypeURL"
        Case pbHlinkTargetTypeEmail: PbHlinkTargetTypeToString = "pbHlinkTargetTypeEmail"
        Case pbHlinkTargetTypeFirstPage: PbHlinkTargetTypeToString = "pbHlinkTargetTypeFirstPage"
        Case pbHlinkTargetTypeLastPage: PbHlinkTargetTypeToString = "pbHlinkTargetTypeLastPage"
        Case pbHlinkTargetTypeNextPage: PbHlinkTargetTypeToString = "pbHlinkTargetTypeNextPage"
        Case pbHlinkTargetTypePreviousPage: PbHlinkTargetTypeToString = "pbHlinkTargetTypePreviousPage"
        Case pbHlinkTargetTypePageID: PbHlinkTargetTypeToString = "pbHlinkTargetTypePageID"
        Case pbHlinkTargetTypePersonalized: PbHlinkTargetTypeToString = "pbHlinkTargetTypePersonalized"
    End Select
End Function
