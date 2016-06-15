Attribute VB_Name = "wMsoFlipCmd"
Function MsoFlipCmdFromString(value As String) As MsoFlipCmd
    If IsNumeric(value) Then
        MsoFlipCmdFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFlipHorizontal": MsoFlipCmdFromString = msoFlipHorizontal
        Case "msoFlipVertical": MsoFlipCmdFromString = msoFlipVertical
    End Select
End Function

Function MsoFlipCmdToString(value As MsoFlipCmd) As String
    Select Case value
        Case msoFlipHorizontal: MsoFlipCmdToString = "msoFlipHorizontal"
        Case msoFlipVertical: MsoFlipCmdToString = "msoFlipVertical"
    End Select
End Function
