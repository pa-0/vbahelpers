Attribute VB_Name = "wMsoZOrderCmd"
Function MsoZOrderCmdFromString(value As String) As MsoZOrderCmd
    If IsNumeric(value) Then
        MsoZOrderCmdFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBringToFront": MsoZOrderCmdFromString = msoBringToFront
        Case "msoSendToBack": MsoZOrderCmdFromString = msoSendToBack
        Case "msoBringForward": MsoZOrderCmdFromString = msoBringForward
        Case "msoSendBackward": MsoZOrderCmdFromString = msoSendBackward
        Case "msoBringInFrontOfText": MsoZOrderCmdFromString = msoBringInFrontOfText
        Case "msoSendBehindText": MsoZOrderCmdFromString = msoSendBehindText
    End Select
End Function

Function MsoZOrderCmdToString(value As MsoZOrderCmd) As String
    Select Case value
        Case msoBringToFront: MsoZOrderCmdToString = "msoBringToFront"
        Case msoSendToBack: MsoZOrderCmdToString = "msoSendToBack"
        Case msoBringForward: MsoZOrderCmdToString = "msoBringForward"
        Case msoSendBackward: MsoZOrderCmdToString = "msoSendBackward"
        Case msoBringInFrontOfText: MsoZOrderCmdToString = "msoBringInFrontOfText"
        Case msoSendBehindText: MsoZOrderCmdToString = "msoSendBehindText"
    End Select
End Function
