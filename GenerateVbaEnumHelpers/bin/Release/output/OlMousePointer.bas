Attribute VB_Name = "wOlMousePointer"
Function OlMousePointerFromString(value As String) As OlMousePointer
    If IsNumeric(value) Then
        OlMousePointerFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMousePointerDefault": OlMousePointerFromString = olMousePointerDefault
        Case "olMousePointerArrow": OlMousePointerFromString = olMousePointerArrow
        Case "olMousePointerCross": OlMousePointerFromString = olMousePointerCross
        Case "olMousePointerIBeam": OlMousePointerFromString = olMousePointerIBeam
        Case "olMousePointerSizeNESW": OlMousePointerFromString = olMousePointerSizeNESW
        Case "olMousePointerSizeNS": OlMousePointerFromString = olMousePointerSizeNS
        Case "olMousePointerSizeNWSE": OlMousePointerFromString = olMousePointerSizeNWSE
        Case "olMousePointerSizeWE": OlMousePointerFromString = olMousePointerSizeWE
        Case "olMousePointerUpArrow": OlMousePointerFromString = olMousePointerUpArrow
        Case "olMousePointerHourGlass": OlMousePointerFromString = olMousePointerHourGlass
        Case "olMousePointerNoDrop": OlMousePointerFromString = olMousePointerNoDrop
        Case "olMousePointerAppStarting": OlMousePointerFromString = olMousePointerAppStarting
        Case "olMousePointerHelp": OlMousePointerFromString = olMousePointerHelp
        Case "olMousePointerSizeAll": OlMousePointerFromString = olMousePointerSizeAll
        Case "olMousePointerCustom": OlMousePointerFromString = olMousePointerCustom
    End Select
End Function

Function OlMousePointerToString(value As OlMousePointer) As String
    Select Case value
        Case olMousePointerDefault: OlMousePointerToString = "olMousePointerDefault"
        Case olMousePointerArrow: OlMousePointerToString = "olMousePointerArrow"
        Case olMousePointerCross: OlMousePointerToString = "olMousePointerCross"
        Case olMousePointerIBeam: OlMousePointerToString = "olMousePointerIBeam"
        Case olMousePointerSizeNESW: OlMousePointerToString = "olMousePointerSizeNESW"
        Case olMousePointerSizeNS: OlMousePointerToString = "olMousePointerSizeNS"
        Case olMousePointerSizeNWSE: OlMousePointerToString = "olMousePointerSizeNWSE"
        Case olMousePointerSizeWE: OlMousePointerToString = "olMousePointerSizeWE"
        Case olMousePointerUpArrow: OlMousePointerToString = "olMousePointerUpArrow"
        Case olMousePointerHourGlass: OlMousePointerToString = "olMousePointerHourGlass"
        Case olMousePointerNoDrop: OlMousePointerToString = "olMousePointerNoDrop"
        Case olMousePointerAppStarting: OlMousePointerToString = "olMousePointerAppStarting"
        Case olMousePointerHelp: OlMousePointerToString = "olMousePointerHelp"
        Case olMousePointerSizeAll: OlMousePointerToString = "olMousePointerSizeAll"
        Case olMousePointerCustom: OlMousePointerToString = "olMousePointerCustom"
    End Select
End Function
