Attribute VB_Name = "wXlEnableSelection"
Function XlEnableSelectionFromString(value As String) As XlEnableSelection
    If IsNumeric(value) Then
        XlEnableSelectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNoRestrictions": XlEnableSelectionFromString = xlNoRestrictions
        Case "xlUnlockedCells": XlEnableSelectionFromString = xlUnlockedCells
        Case "xlNoSelection": XlEnableSelectionFromString = xlNoSelection
    End Select
End Function

Function XlEnableSelectionToString(value As XlEnableSelection) As String
    Select Case value
        Case xlNoRestrictions: XlEnableSelectionToString = "xlNoRestrictions"
        Case xlUnlockedCells: XlEnableSelectionToString = "xlUnlockedCells"
        Case xlNoSelection: XlEnableSelectionToString = "xlNoSelection"
    End Select
End Function
