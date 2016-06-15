Attribute VB_Name = "wMsoExtraInfoMethod"
Function MsoExtraInfoMethodFromString(value As String) As MsoExtraInfoMethod
    If IsNumeric(value) Then
        MsoExtraInfoMethodFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoMethodGet": MsoExtraInfoMethodFromString = msoMethodGet
        Case "msoMethodPost": MsoExtraInfoMethodFromString = msoMethodPost
    End Select
End Function

Function MsoExtraInfoMethodToString(value As MsoExtraInfoMethod) As String
    Select Case value
        Case msoMethodGet: MsoExtraInfoMethodToString = "msoMethodGet"
        Case msoMethodPost: MsoExtraInfoMethodToString = "msoMethodPost"
    End Select
End Function
