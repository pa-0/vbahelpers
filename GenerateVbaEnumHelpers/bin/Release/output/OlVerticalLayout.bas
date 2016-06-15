Attribute VB_Name = "wOlVerticalLayout"
Function OlVerticalLayoutFromString(value As String) As OlVerticalLayout
    If IsNumeric(value) Then
        OlVerticalLayoutFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olVerticalLayoutAlignTop": OlVerticalLayoutFromString = olVerticalLayoutAlignTop
        Case "olVerticalLayoutAlignMiddle": OlVerticalLayoutFromString = olVerticalLayoutAlignMiddle
        Case "olVerticalLayoutAlignBottom": OlVerticalLayoutFromString = olVerticalLayoutAlignBottom
        Case "olVerticalLayoutGrow": OlVerticalLayoutFromString = olVerticalLayoutGrow
    End Select
End Function

Function OlVerticalLayoutToString(value As OlVerticalLayout) As String
    Select Case value
        Case olVerticalLayoutAlignTop: OlVerticalLayoutToString = "olVerticalLayoutAlignTop"
        Case olVerticalLayoutAlignMiddle: OlVerticalLayoutToString = "olVerticalLayoutAlignMiddle"
        Case olVerticalLayoutAlignBottom: OlVerticalLayoutToString = "olVerticalLayoutAlignBottom"
        Case olVerticalLayoutGrow: OlVerticalLayoutToString = "olVerticalLayoutGrow"
    End Select
End Function
