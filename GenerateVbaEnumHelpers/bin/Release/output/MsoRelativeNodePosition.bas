Attribute VB_Name = "wMsoRelativeNodePosition"
Function MsoRelativeNodePositionFromString(value As String) As MsoRelativeNodePosition
    If IsNumeric(value) Then
        MsoRelativeNodePositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBeforeNode": MsoRelativeNodePositionFromString = msoBeforeNode
        Case "msoAfterNode": MsoRelativeNodePositionFromString = msoAfterNode
        Case "msoBeforeFirstSibling": MsoRelativeNodePositionFromString = msoBeforeFirstSibling
        Case "msoAfterLastSibling": MsoRelativeNodePositionFromString = msoAfterLastSibling
    End Select
End Function

Function MsoRelativeNodePositionToString(value As MsoRelativeNodePosition) As String
    Select Case value
        Case msoBeforeNode: MsoRelativeNodePositionToString = "msoBeforeNode"
        Case msoAfterNode: MsoRelativeNodePositionToString = "msoAfterNode"
        Case msoBeforeFirstSibling: MsoRelativeNodePositionToString = "msoBeforeFirstSibling"
        Case msoAfterLastSibling: MsoRelativeNodePositionToString = "msoAfterLastSibling"
    End Select
End Function
