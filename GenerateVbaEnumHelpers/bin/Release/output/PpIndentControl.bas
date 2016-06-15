Attribute VB_Name = "wPpIndentControl"
Function PpIndentControlFromString(value As String) As PpIndentControl
    If IsNumeric(value) Then
        PpIndentControlFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppIndentReplaceAttr": PpIndentControlFromString = ppIndentReplaceAttr
        Case "ppIndentKeepAttr": PpIndentControlFromString = ppIndentKeepAttr
        Case "ppIndentControlMixed": PpIndentControlFromString = ppIndentControlMixed
    End Select
End Function

Function PpIndentControlToString(value As PpIndentControl) As String
    Select Case value
        Case ppIndentReplaceAttr: PpIndentControlToString = "ppIndentReplaceAttr"
        Case ppIndentKeepAttr: PpIndentControlToString = "ppIndentKeepAttr"
        Case ppIndentControlMixed: PpIndentControlToString = "ppIndentControlMixed"
    End Select
End Function
