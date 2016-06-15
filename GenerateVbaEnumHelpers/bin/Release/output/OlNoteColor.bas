Attribute VB_Name = "wOlNoteColor"
Function OlNoteColorFromString(value As String) As OlNoteColor
    If IsNumeric(value) Then
        OlNoteColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olBlue": OlNoteColorFromString = olBlue
        Case "olGreen": OlNoteColorFromString = olGreen
        Case "olPink": OlNoteColorFromString = olPink
        Case "olYellow": OlNoteColorFromString = olYellow
        Case "olWhite": OlNoteColorFromString = olWhite
    End Select
End Function

Function OlNoteColorToString(value As OlNoteColor) As String
    Select Case value
        Case olBlue: OlNoteColorToString = "olBlue"
        Case olGreen: OlNoteColorToString = "olGreen"
        Case olPink: OlNoteColorToString = "olPink"
        Case olYellow: OlNoteColorToString = "olYellow"
        Case olWhite: OlNoteColorToString = "olWhite"
    End Select
End Function
