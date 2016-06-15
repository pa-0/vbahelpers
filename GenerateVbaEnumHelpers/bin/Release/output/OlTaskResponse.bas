Attribute VB_Name = "wOlTaskResponse"
Function OlTaskResponseFromString(value As String) As OlTaskResponse
    If IsNumeric(value) Then
        OlTaskResponseFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTaskSimple": OlTaskResponseFromString = olTaskSimple
        Case "olTaskAssign": OlTaskResponseFromString = olTaskAssign
        Case "olTaskAccept": OlTaskResponseFromString = olTaskAccept
        Case "olTaskDecline": OlTaskResponseFromString = olTaskDecline
    End Select
End Function

Function OlTaskResponseToString(value As OlTaskResponse) As String
    Select Case value
        Case olTaskSimple: OlTaskResponseToString = "olTaskSimple"
        Case olTaskAssign: OlTaskResponseToString = "olTaskAssign"
        Case olTaskAccept: OlTaskResponseToString = "olTaskAccept"
        Case olTaskDecline: OlTaskResponseToString = "olTaskDecline"
    End Select
End Function
