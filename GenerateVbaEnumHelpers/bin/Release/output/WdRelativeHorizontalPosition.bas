Attribute VB_Name = "wWdRelativeHorizontalPosition"
Function WdRelativeHorizontalPositionFromString(value As String) As WdRelativeHorizontalPosition
    If IsNumeric(value) Then
        WdRelativeHorizontalPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRelativeHorizontalPositionMargin": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionMargin
        Case "wdRelativeHorizontalPositionPage": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionPage
        Case "wdRelativeHorizontalPositionColumn": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionColumn
        Case "wdRelativeHorizontalPositionCharacter": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionCharacter
        Case "wdRelativeHorizontalPositionLeftMarginArea": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionLeftMarginArea
        Case "wdRelativeHorizontalPositionRightMarginArea": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionRightMarginArea
        Case "wdRelativeHorizontalPositionInnerMarginArea": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionInnerMarginArea
        Case "wdRelativeHorizontalPositionOuterMarginArea": WdRelativeHorizontalPositionFromString = wdRelativeHorizontalPositionOuterMarginArea
    End Select
End Function

Function WdRelativeHorizontalPositionToString(value As WdRelativeHorizontalPosition) As String
    Select Case value
        Case wdRelativeHorizontalPositionMargin: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionMargin"
        Case wdRelativeHorizontalPositionPage: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionPage"
        Case wdRelativeHorizontalPositionColumn: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionColumn"
        Case wdRelativeHorizontalPositionCharacter: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionCharacter"
        Case wdRelativeHorizontalPositionLeftMarginArea: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionLeftMarginArea"
        Case wdRelativeHorizontalPositionRightMarginArea: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionRightMarginArea"
        Case wdRelativeHorizontalPositionInnerMarginArea: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionInnerMarginArea"
        Case wdRelativeHorizontalPositionOuterMarginArea: WdRelativeHorizontalPositionToString = "wdRelativeHorizontalPositionOuterMarginArea"
    End Select
End Function
