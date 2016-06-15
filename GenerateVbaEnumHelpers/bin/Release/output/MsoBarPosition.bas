Attribute VB_Name = "wMsoBarPosition"
Function MsoBarPositionFromString(value As String) As MsoBarPosition
    If IsNumeric(value) Then
        MsoBarPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBarLeft": MsoBarPositionFromString = msoBarLeft
        Case "msoBarTop": MsoBarPositionFromString = msoBarTop
        Case "msoBarRight": MsoBarPositionFromString = msoBarRight
        Case "msoBarBottom": MsoBarPositionFromString = msoBarBottom
        Case "msoBarFloating": MsoBarPositionFromString = msoBarFloating
        Case "msoBarPopup": MsoBarPositionFromString = msoBarPopup
        Case "msoBarMenuBar": MsoBarPositionFromString = msoBarMenuBar
    End Select
End Function

Function MsoBarPositionToString(value As MsoBarPosition) As String
    Select Case value
        Case msoBarLeft: MsoBarPositionToString = "msoBarLeft"
        Case msoBarTop: MsoBarPositionToString = "msoBarTop"
        Case msoBarRight: MsoBarPositionToString = "msoBarRight"
        Case msoBarBottom: MsoBarPositionToString = "msoBarBottom"
        Case msoBarFloating: MsoBarPositionToString = "msoBarFloating"
        Case msoBarPopup: MsoBarPositionToString = "msoBarPopup"
        Case msoBarMenuBar: MsoBarPositionToString = "msoBarMenuBar"
    End Select
End Function
