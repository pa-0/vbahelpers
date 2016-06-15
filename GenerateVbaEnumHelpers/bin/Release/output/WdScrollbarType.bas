Attribute VB_Name = "wWdScrollbarType"
Function WdScrollbarTypeFromString(value As String) As WdScrollbarType
    If IsNumeric(value) Then
        WdScrollbarTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdScrollbarTypeAuto": WdScrollbarTypeFromString = wdScrollbarTypeAuto
        Case "wdScrollbarTypeYes": WdScrollbarTypeFromString = wdScrollbarTypeYes
        Case "wdScrollbarTypeNo": WdScrollbarTypeFromString = wdScrollbarTypeNo
    End Select
End Function

Function WdScrollbarTypeToString(value As WdScrollbarType) As String
    Select Case value
        Case wdScrollbarTypeAuto: WdScrollbarTypeToString = "wdScrollbarTypeAuto"
        Case wdScrollbarTypeYes: WdScrollbarTypeToString = "wdScrollbarTypeYes"
        Case wdScrollbarTypeNo: WdScrollbarTypeToString = "wdScrollbarTypeNo"
    End Select
End Function
