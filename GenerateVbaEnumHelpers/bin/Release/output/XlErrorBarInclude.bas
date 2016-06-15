Attribute VB_Name = "wXlErrorBarInclude"
Function XlErrorBarIncludeFromString(value As String) As XlErrorBarInclude
    If IsNumeric(value) Then
        XlErrorBarIncludeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlErrorBarIncludeBoth": XlErrorBarIncludeFromString = xlErrorBarIncludeBoth
        Case "xlErrorBarIncludePlusValues": XlErrorBarIncludeFromString = xlErrorBarIncludePlusValues
        Case "xlErrorBarIncludeMinusValues": XlErrorBarIncludeFromString = xlErrorBarIncludeMinusValues
        Case "xlErrorBarIncludeNone": XlErrorBarIncludeFromString = xlErrorBarIncludeNone
    End Select
End Function

Function XlErrorBarIncludeToString(value As XlErrorBarInclude) As String
    Select Case value
        Case xlErrorBarIncludeBoth: XlErrorBarIncludeToString = "xlErrorBarIncludeBoth"
        Case xlErrorBarIncludePlusValues: XlErrorBarIncludeToString = "xlErrorBarIncludePlusValues"
        Case xlErrorBarIncludeMinusValues: XlErrorBarIncludeToString = "xlErrorBarIncludeMinusValues"
        Case xlErrorBarIncludeNone: XlErrorBarIncludeToString = "xlErrorBarIncludeNone"
    End Select
End Function
