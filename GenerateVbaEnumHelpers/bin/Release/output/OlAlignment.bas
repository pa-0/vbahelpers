Attribute VB_Name = "wOlAlignment"
Function OlAlignmentFromString(value As String) As OlAlignment
    If IsNumeric(value) Then
        OlAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAlignmentLeft": OlAlignmentFromString = olAlignmentLeft
        Case "olAlignmentRight": OlAlignmentFromString = olAlignmentRight
    End Select
End Function

Function OlAlignmentToString(value As OlAlignment) As String
    Select Case value
        Case olAlignmentLeft: OlAlignmentToString = "olAlignmentLeft"
        Case olAlignmentRight: OlAlignmentToString = "olAlignmentRight"
    End Select
End Function
