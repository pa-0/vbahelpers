Attribute VB_Name = "wXlMousePointer"
Function XlMousePointerFromString(value As String) As XlMousePointer
    If IsNumeric(value) Then
        XlMousePointerFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNorthwestArrow": XlMousePointerFromString = xlNorthwestArrow
        Case "xlWait": XlMousePointerFromString = xlWait
        Case "xlIBeam": XlMousePointerFromString = xlIBeam
        Case "xlDefault": XlMousePointerFromString = xlDefault
    End Select
End Function

Function XlMousePointerToString(value As XlMousePointer) As String
    Select Case value
        Case xlNorthwestArrow: XlMousePointerToString = "xlNorthwestArrow"
        Case xlWait: XlMousePointerToString = "xlWait"
        Case xlIBeam: XlMousePointerToString = "xlIBeam"
        Case xlDefault: XlMousePointerToString = "xlDefault"
    End Select
End Function
