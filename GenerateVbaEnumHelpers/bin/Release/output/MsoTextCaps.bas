Attribute VB_Name = "wMsoTextCaps"
Function MsoTextCapsFromString(value As String) As MsoTextCaps
    If IsNumeric(value) Then
        MsoTextCapsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoNoCaps": MsoTextCapsFromString = msoNoCaps
        Case "msoSmallCaps": MsoTextCapsFromString = msoSmallCaps
        Case "msoAllCaps": MsoTextCapsFromString = msoAllCaps
        Case "msoCapsMixed": MsoTextCapsFromString = msoCapsMixed
    End Select
End Function

Function MsoTextCapsToString(value As MsoTextCaps) As String
    Select Case value
        Case msoNoCaps: MsoTextCapsToString = "msoNoCaps"
        Case msoSmallCaps: MsoTextCapsToString = "msoSmallCaps"
        Case msoAllCaps: MsoTextCapsToString = "msoAllCaps"
        Case msoCapsMixed: MsoTextCapsToString = "msoCapsMixed"
    End Select
End Function
