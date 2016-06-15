Attribute VB_Name = "wWdNumberStyleWordBasicBiDi"
Function WdNumberStyleWordBasicBiDiFromString(value As String) As WdNumberStyleWordBasicBiDi
    If IsNumeric(value) Then
        WdNumberStyleWordBasicBiDiFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCaptionNumberStyleBidiLetter1": WdNumberStyleWordBasicBiDiFromString = wdCaptionNumberStyleBidiLetter1
        Case "wdListNumberStyleBidi1": WdNumberStyleWordBasicBiDiFromString = wdListNumberStyleBidi1
        Case "wdPageNumberStyleBidiLetter1": WdNumberStyleWordBasicBiDiFromString = wdPageNumberStyleBidiLetter1
        Case "wdNoteNumberStyleBidiLetter1": WdNumberStyleWordBasicBiDiFromString = wdNoteNumberStyleBidiLetter1
        Case "wdCaptionNumberStyleBidiLetter2": WdNumberStyleWordBasicBiDiFromString = wdCaptionNumberStyleBidiLetter2
        Case "wdListNumberStyleBidi2": WdNumberStyleWordBasicBiDiFromString = wdListNumberStyleBidi2
        Case "wdNoteNumberStyleBidiLetter2": WdNumberStyleWordBasicBiDiFromString = wdNoteNumberStyleBidiLetter2
        Case "wdPageNumberStyleBidiLetter2": WdNumberStyleWordBasicBiDiFromString = wdPageNumberStyleBidiLetter2
    End Select
End Function

Function WdNumberStyleWordBasicBiDiToString(value As WdNumberStyleWordBasicBiDi) As String
    Select Case value
        Case wdCaptionNumberStyleBidiLetter1: WdNumberStyleWordBasicBiDiToString = "wdCaptionNumberStyleBidiLetter1"
        Case wdListNumberStyleBidi1: WdNumberStyleWordBasicBiDiToString = "wdListNumberStyleBidi1"
        Case wdPageNumberStyleBidiLetter1: WdNumberStyleWordBasicBiDiToString = "wdPageNumberStyleBidiLetter1"
        Case wdNoteNumberStyleBidiLetter1: WdNumberStyleWordBasicBiDiToString = "wdNoteNumberStyleBidiLetter1"
        Case wdCaptionNumberStyleBidiLetter2: WdNumberStyleWordBasicBiDiToString = "wdCaptionNumberStyleBidiLetter2"
        Case wdListNumberStyleBidi2: WdNumberStyleWordBasicBiDiToString = "wdListNumberStyleBidi2"
        Case wdNoteNumberStyleBidiLetter2: WdNumberStyleWordBasicBiDiToString = "wdNoteNumberStyleBidiLetter2"
        Case wdPageNumberStyleBidiLetter2: WdNumberStyleWordBasicBiDiToString = "wdPageNumberStyleBidiLetter2"
    End Select
End Function
