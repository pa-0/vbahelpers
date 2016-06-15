Attribute VB_Name = "wOlPageType"
Function OlPageTypeFromString(value As String) As OlPageType
    If IsNumeric(value) Then
        OlPageTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olPageTypePlanner": OlPageTypeFromString = olPageTypePlanner
        Case "olPageTypeTracker": OlPageTypeFromString = olPageTypeTracker
    End Select
End Function

Function OlPageTypeToString(value As OlPageType) As String
    Select Case value
        Case olPageTypePlanner: OlPageTypeToString = "olPageTypePlanner"
        Case olPageTypeTracker: OlPageTypeToString = "olPageTypeTracker"
    End Select
End Function
