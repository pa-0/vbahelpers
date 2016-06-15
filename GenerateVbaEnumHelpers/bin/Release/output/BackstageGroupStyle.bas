Attribute VB_Name = "wBackstageGroupStyle"
Function BackstageGroupStyleFromString(value As String) As BackstageGroupStyle
    If IsNumeric(value) Then
        BackstageGroupStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "BackstageGroupStyleNormal": BackstageGroupStyleFromString = BackstageGroupStyleNormal
        Case "BackstageGroupStyleWarning": BackstageGroupStyleFromString = BackstageGroupStyleWarning
        Case "BackstageGroupStyleError": BackstageGroupStyleFromString = BackstageGroupStyleError
    End Select
End Function

Function BackstageGroupStyleToString(value As BackstageGroupStyle) As String
    Select Case value
        Case BackstageGroupStyleNormal: BackstageGroupStyleToString = "BackstageGroupStyleNormal"
        Case BackstageGroupStyleWarning: BackstageGroupStyleToString = "BackstageGroupStyleWarning"
        Case BackstageGroupStyleError: BackstageGroupStyleToString = "BackstageGroupStyleError"
    End Select
End Function
