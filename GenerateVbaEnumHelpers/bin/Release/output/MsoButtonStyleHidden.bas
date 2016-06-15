Attribute VB_Name = "wMsoButtonStyleHidden"
Function MsoButtonStyleHiddenFromString(value As String) As MsoButtonStyleHidden
    If IsNumeric(value) Then
        MsoButtonStyleHiddenFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoButtonWrapText": MsoButtonStyleHiddenFromString = msoButtonWrapText
        Case "msoButtonTextBelow": MsoButtonStyleHiddenFromString = msoButtonTextBelow
    End Select
End Function

Function MsoButtonStyleHiddenToString(value As MsoButtonStyleHidden) As String
    Select Case value
        Case msoButtonWrapText: MsoButtonStyleHiddenToString = "msoButtonWrapText"
        Case msoButtonTextBelow: MsoButtonStyleHiddenToString = "msoButtonTextBelow"
    End Select
End Function
