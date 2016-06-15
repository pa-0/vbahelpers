Attribute VB_Name = "wMsoTriState"
Function MsoTriStateFromString(value As String) As MsoTriState
    If IsNumeric(value) Then
        MsoTriStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFalse": MsoTriStateFromString = msoFalse
        Case "msoCTrue": MsoTriStateFromString = msoCTrue
        Case "msoTriStateToggle": MsoTriStateFromString = msoTriStateToggle
        Case "msoTriStateMixed": MsoTriStateFromString = msoTriStateMixed
        Case "msoTrue": MsoTriStateFromString = msoTrue
    End Select
End Function

Function MsoTriStateToString(value As MsoTriState) As String
    Select Case value
        Case msoFalse: MsoTriStateToString = "msoFalse"
        Case msoCTrue: MsoTriStateToString = "msoCTrue"
        Case msoTriStateToggle: MsoTriStateToString = "msoTriStateToggle"
        Case msoTriStateMixed: MsoTriStateToString = "msoTriStateMixed"
        Case msoTrue: MsoTriStateToString = "msoTrue"
    End Select
End Function
