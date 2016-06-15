Attribute VB_Name = "wMsoFileFindOptions"
Function MsoFileFindOptionsFromString(value As String) As MsoFileFindOptions
    If IsNumeric(value) Then
        MsoFileFindOptionsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoOptionsNew": MsoFileFindOptionsFromString = msoOptionsNew
        Case "msoOptionsAdd": MsoFileFindOptionsFromString = msoOptionsAdd
        Case "msoOptionsWithin": MsoFileFindOptionsFromString = msoOptionsWithin
    End Select
End Function

Function MsoFileFindOptionsToString(value As MsoFileFindOptions) As String
    Select Case value
        Case msoOptionsNew: MsoFileFindOptionsToString = "msoOptionsNew"
        Case msoOptionsAdd: MsoFileFindOptionsToString = "msoOptionsAdd"
        Case msoOptionsWithin: MsoFileFindOptionsToString = "msoOptionsWithin"
    End Select
End Function
