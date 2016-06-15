Attribute VB_Name = "wWdWrapTypeMerged"
Function WdWrapTypeMergedFromString(value As String) As WdWrapTypeMerged
    If IsNumeric(value) Then
        WdWrapTypeMergedFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWrapMergeInline": WdWrapTypeMergedFromString = wdWrapMergeInline
        Case "wdWrapMergeSquare": WdWrapTypeMergedFromString = wdWrapMergeSquare
        Case "wdWrapMergeTight": WdWrapTypeMergedFromString = wdWrapMergeTight
        Case "wdWrapMergeBehind": WdWrapTypeMergedFromString = wdWrapMergeBehind
        Case "wdWrapMergeFront": WdWrapTypeMergedFromString = wdWrapMergeFront
        Case "wdWrapMergeThrough": WdWrapTypeMergedFromString = wdWrapMergeThrough
        Case "wdWrapMergeTopBottom": WdWrapTypeMergedFromString = wdWrapMergeTopBottom
    End Select
End Function

Function WdWrapTypeMergedToString(value As WdWrapTypeMerged) As String
    Select Case value
        Case wdWrapMergeInline: WdWrapTypeMergedToString = "wdWrapMergeInline"
        Case wdWrapMergeSquare: WdWrapTypeMergedToString = "wdWrapMergeSquare"
        Case wdWrapMergeTight: WdWrapTypeMergedToString = "wdWrapMergeTight"
        Case wdWrapMergeBehind: WdWrapTypeMergedToString = "wdWrapMergeBehind"
        Case wdWrapMergeFront: WdWrapTypeMergedToString = "wdWrapMergeFront"
        Case wdWrapMergeThrough: WdWrapTypeMergedToString = "wdWrapMergeThrough"
        Case wdWrapMergeTopBottom: WdWrapTypeMergedToString = "wdWrapMergeTopBottom"
    End Select
End Function
