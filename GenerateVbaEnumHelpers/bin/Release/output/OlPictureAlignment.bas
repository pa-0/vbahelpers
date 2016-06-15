Attribute VB_Name = "wOlPictureAlignment"
Function OlPictureAlignmentFromString(value As String) As OlPictureAlignment
    If IsNumeric(value) Then
        OlPictureAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olPictureAlignmentLeft": OlPictureAlignmentFromString = olPictureAlignmentLeft
        Case "olPictureAlignmentTop": OlPictureAlignmentFromString = olPictureAlignmentTop
    End Select
End Function

Function OlPictureAlignmentToString(value As OlPictureAlignment) As String
    Select Case value
        Case olPictureAlignmentLeft: OlPictureAlignmentToString = "olPictureAlignmentLeft"
        Case olPictureAlignmentTop: OlPictureAlignmentToString = "olPictureAlignmentTop"
    End Select
End Function
