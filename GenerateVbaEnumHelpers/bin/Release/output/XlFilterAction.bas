Attribute VB_Name = "wXlFilterAction"
Function XlFilterActionFromString(value As String) As XlFilterAction
    If IsNumeric(value) Then
        XlFilterActionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFilterInPlace": XlFilterActionFromString = xlFilterInPlace
        Case "xlFilterCopy": XlFilterActionFromString = xlFilterCopy
    End Select
End Function

Function XlFilterActionToString(value As XlFilterAction) As String
    Select Case value
        Case xlFilterInPlace: XlFilterActionToString = "xlFilterInPlace"
        Case xlFilterCopy: XlFilterActionToString = "xlFilterCopy"
    End Select
End Function
