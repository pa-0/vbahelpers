Attribute VB_Name = "wPbSubmitDataRetrievalMethodType"
Function PbSubmitDataRetrievalMethodTypeFromString(value As String) As PbSubmitDataRetrievalMethodType
    If IsNumeric(value) Then
        PbSubmitDataRetrievalMethodTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbSubmitDataRetrievalSaveOnServer": PbSubmitDataRetrievalMethodTypeFromString = pbSubmitDataRetrievalSaveOnServer
        Case "pbSubmitDataRetrievalEmail": PbSubmitDataRetrievalMethodTypeFromString = pbSubmitDataRetrievalEmail
        Case "pbSubmitDataRetrievalProgram": PbSubmitDataRetrievalMethodTypeFromString = pbSubmitDataRetrievalProgram
    End Select
End Function

Function PbSubmitDataRetrievalMethodTypeToString(value As PbSubmitDataRetrievalMethodType) As String
    Select Case value
        Case pbSubmitDataRetrievalSaveOnServer: PbSubmitDataRetrievalMethodTypeToString = "pbSubmitDataRetrievalSaveOnServer"
        Case pbSubmitDataRetrievalEmail: PbSubmitDataRetrievalMethodTypeToString = "pbSubmitDataRetrievalEmail"
        Case pbSubmitDataRetrievalProgram: PbSubmitDataRetrievalMethodTypeToString = "pbSubmitDataRetrievalProgram"
    End Select
End Function
