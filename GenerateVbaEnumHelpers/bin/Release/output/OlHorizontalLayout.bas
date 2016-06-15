Attribute VB_Name = "wOlHorizontalLayout"
Function OlHorizontalLayoutFromString(value As String) As OlHorizontalLayout
    If IsNumeric(value) Then
        OlHorizontalLayoutFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olHorizontalLayoutAlignLeft": OlHorizontalLayoutFromString = olHorizontalLayoutAlignLeft
        Case "olHorizontalLayoutAlignCenter": OlHorizontalLayoutFromString = olHorizontalLayoutAlignCenter
        Case "olHorizontalLayoutAlignRight": OlHorizontalLayoutFromString = olHorizontalLayoutAlignRight
        Case "olHorizontalLayoutGrow": OlHorizontalLayoutFromString = olHorizontalLayoutGrow
    End Select
End Function

Function OlHorizontalLayoutToString(value As OlHorizontalLayout) As String
    Select Case value
        Case olHorizontalLayoutAlignLeft: OlHorizontalLayoutToString = "olHorizontalLayoutAlignLeft"
        Case olHorizontalLayoutAlignCenter: OlHorizontalLayoutToString = "olHorizontalLayoutAlignCenter"
        Case olHorizontalLayoutAlignRight: OlHorizontalLayoutToString = "olHorizontalLayoutAlignRight"
        Case olHorizontalLayoutGrow: OlHorizontalLayoutToString = "olHorizontalLayoutGrow"
    End Select
End Function
