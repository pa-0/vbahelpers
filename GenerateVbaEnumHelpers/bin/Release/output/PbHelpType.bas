Attribute VB_Name = "wPbHelpType"
Function PbHelpTypeFromString(value As String) As PbHelpType
    If IsNumeric(value) Then
        PbHelpTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbHelp": PbHelpTypeFromString = pbHelp
        Case "pbHelpActiveWindow": PbHelpTypeFromString = pbHelpActiveWindow
        Case "pbHelpPSSHelp": PbHelpTypeFromString = pbHelpPSSHelp
    End Select
End Function

Function PbHelpTypeToString(value As PbHelpType) As String
    Select Case value
        Case pbHelp: PbHelpTypeToString = "pbHelp"
        Case pbHelpActiveWindow: PbHelpTypeToString = "pbHelpActiveWindow"
        Case pbHelpPSSHelp: PbHelpTypeToString = "pbHelpPSSHelp"
    End Select
End Function
