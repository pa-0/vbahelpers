Attribute VB_Name = "wOlIconViewType"
Function OlIconViewTypeFromString(value As String) As OlIconViewType
    If IsNumeric(value) Then
        OlIconViewTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olIconViewLarge": OlIconViewTypeFromString = olIconViewLarge
        Case "olIconViewSmall": OlIconViewTypeFromString = olIconViewSmall
        Case "olIconViewList": OlIconViewTypeFromString = olIconViewList
    End Select
End Function

Function OlIconViewTypeToString(value As OlIconViewType) As String
    Select Case value
        Case olIconViewLarge: OlIconViewTypeToString = "olIconViewLarge"
        Case olIconViewSmall: OlIconViewTypeToString = "olIconViewSmall"
        Case olIconViewList: OlIconViewTypeToString = "olIconViewList"
    End Select
End Function
