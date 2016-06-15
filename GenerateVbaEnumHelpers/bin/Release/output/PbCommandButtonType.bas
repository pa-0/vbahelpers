Attribute VB_Name = "wPbCommandButtonType"
Function PbCommandButtonTypeFromString(value As String) As PbCommandButtonType
    If IsNumeric(value) Then
        PbCommandButtonTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbCommandButtonSubmit": PbCommandButtonTypeFromString = pbCommandButtonSubmit
        Case "pbCommandButtonReset": PbCommandButtonTypeFromString = pbCommandButtonReset
    End Select
End Function

Function PbCommandButtonTypeToString(value As PbCommandButtonType) As String
    Select Case value
        Case pbCommandButtonSubmit: PbCommandButtonTypeToString = "pbCommandButtonSubmit"
        Case pbCommandButtonReset: PbCommandButtonTypeToString = "pbCommandButtonReset"
    End Select
End Function
