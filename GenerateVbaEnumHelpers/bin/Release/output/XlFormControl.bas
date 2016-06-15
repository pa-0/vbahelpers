Attribute VB_Name = "wXlFormControl"
Function XlFormControlFromString(value As String) As XlFormControl
    If IsNumeric(value) Then
        XlFormControlFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlButtonControl": XlFormControlFromString = xlButtonControl
        Case "xlCheckBox": XlFormControlFromString = xlCheckBox
        Case "xlDropDown": XlFormControlFromString = xlDropDown
        Case "xlEditBox": XlFormControlFromString = xlEditBox
        Case "xlGroupBox": XlFormControlFromString = xlGroupBox
        Case "xlLabel": XlFormControlFromString = xlLabel
        Case "xlListBox": XlFormControlFromString = xlListBox
        Case "xlOptionButton": XlFormControlFromString = xlOptionButton
        Case "xlScrollBar": XlFormControlFromString = xlScrollBar
        Case "xlSpinner": XlFormControlFromString = xlSpinner
    End Select
End Function

Function XlFormControlToString(value As XlFormControl) As String
    Select Case value
        Case xlButtonControl: XlFormControlToString = "xlButtonControl"
        Case xlCheckBox: XlFormControlToString = "xlCheckBox"
        Case xlDropDown: XlFormControlToString = "xlDropDown"
        Case xlEditBox: XlFormControlToString = "xlEditBox"
        Case xlGroupBox: XlFormControlToString = "xlGroupBox"
        Case xlLabel: XlFormControlToString = "xlLabel"
        Case xlListBox: XlFormControlToString = "xlListBox"
        Case xlOptionButton: XlFormControlToString = "xlOptionButton"
        Case xlScrollBar: XlFormControlToString = "xlScrollBar"
        Case xlSpinner: XlFormControlToString = "xlSpinner"
    End Select
End Function
