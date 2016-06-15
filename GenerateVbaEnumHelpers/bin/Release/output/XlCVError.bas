Attribute VB_Name = "wXlCVError"
Function XlCVErrorFromString(value As String) As XlCVError
    If IsNumeric(value) Then
        XlCVErrorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlErrNull": XlCVErrorFromString = xlErrNull
        Case "xlErrDiv0": XlCVErrorFromString = xlErrDiv0
        Case "xlErrValue": XlCVErrorFromString = xlErrValue
        Case "xlErrRef": XlCVErrorFromString = xlErrRef
        Case "xlErrName": XlCVErrorFromString = xlErrName
        Case "xlErrNum": XlCVErrorFromString = xlErrNum
        Case "xlErrNA": XlCVErrorFromString = xlErrNA
    End Select
End Function

Function XlCVErrorToString(value As XlCVError) As String
    Select Case value
        Case xlErrNull: XlCVErrorToString = "xlErrNull"
        Case xlErrDiv0: XlCVErrorToString = "xlErrDiv0"
        Case xlErrValue: XlCVErrorToString = "xlErrValue"
        Case xlErrRef: XlCVErrorToString = "xlErrRef"
        Case xlErrName: XlCVErrorToString = "xlErrName"
        Case xlErrNum: XlCVErrorToString = "xlErrNum"
        Case xlErrNA: XlCVErrorToString = "xlErrNA"
    End Select
End Function
