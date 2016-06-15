Attribute VB_Name = "wPbPublicationType"
Function PbPublicationTypeFromString(value As String) As PbPublicationType
    If IsNumeric(value) Then
        PbPublicationTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTypePrint": PbPublicationTypeFromString = pbTypePrint
        Case "pbTypeWeb": PbPublicationTypeFromString = pbTypeWeb
    End Select
End Function

Function PbPublicationTypeToString(value As PbPublicationType) As String
    Select Case value
        Case pbTypePrint: PbPublicationTypeToString = "pbTypePrint"
        Case pbTypeWeb: PbPublicationTypeToString = "pbTypeWeb"
    End Select
End Function
