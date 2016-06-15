Attribute VB_Name = "wWdWrapSideType"
Function WdWrapSideTypeFromString(value As String) As WdWrapSideType
    If IsNumeric(value) Then
        WdWrapSideTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWrapBoth": WdWrapSideTypeFromString = wdWrapBoth
        Case "wdWrapLeft": WdWrapSideTypeFromString = wdWrapLeft
        Case "wdWrapRight": WdWrapSideTypeFromString = wdWrapRight
        Case "wdWrapLargest": WdWrapSideTypeFromString = wdWrapLargest
    End Select
End Function

Function WdWrapSideTypeToString(value As WdWrapSideType) As String
    Select Case value
        Case wdWrapBoth: WdWrapSideTypeToString = "wdWrapBoth"
        Case wdWrapLeft: WdWrapSideTypeToString = "wdWrapLeft"
        Case wdWrapRight: WdWrapSideTypeToString = "wdWrapRight"
        Case wdWrapLargest: WdWrapSideTypeToString = "wdWrapLargest"
    End Select
End Function
