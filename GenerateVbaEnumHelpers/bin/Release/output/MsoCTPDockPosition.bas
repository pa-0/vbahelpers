Attribute VB_Name = "wMsoCTPDockPosition"
Function MsoCTPDockPositionFromString(value As String) As MsoCTPDockPosition
    If IsNumeric(value) Then
        MsoCTPDockPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCTPDockPositionLeft": MsoCTPDockPositionFromString = msoCTPDockPositionLeft
        Case "msoCTPDockPositionTop": MsoCTPDockPositionFromString = msoCTPDockPositionTop
        Case "msoCTPDockPositionRight": MsoCTPDockPositionFromString = msoCTPDockPositionRight
        Case "msoCTPDockPositionBottom": MsoCTPDockPositionFromString = msoCTPDockPositionBottom
        Case "msoCTPDockPositionFloating": MsoCTPDockPositionFromString = msoCTPDockPositionFloating
    End Select
End Function

Function MsoCTPDockPositionToString(value As MsoCTPDockPosition) As String
    Select Case value
        Case msoCTPDockPositionLeft: MsoCTPDockPositionToString = "msoCTPDockPositionLeft"
        Case msoCTPDockPositionTop: MsoCTPDockPositionToString = "msoCTPDockPositionTop"
        Case msoCTPDockPositionRight: MsoCTPDockPositionToString = "msoCTPDockPositionRight"
        Case msoCTPDockPositionBottom: MsoCTPDockPositionToString = "msoCTPDockPositionBottom"
        Case msoCTPDockPositionFloating: MsoCTPDockPositionToString = "msoCTPDockPositionFloating"
    End Select
End Function
