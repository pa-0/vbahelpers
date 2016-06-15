Attribute VB_Name = "wPbTextAutoFitType"
Function PbTextAutoFitTypeFromString(value As String) As PbTextAutoFitType
    If IsNumeric(value) Then
        PbTextAutoFitTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTextAutoFitNone": PbTextAutoFitTypeFromString = pbTextAutoFitNone
        Case "pbTextAutoFitShrinkOnOverflow": PbTextAutoFitTypeFromString = pbTextAutoFitShrinkOnOverflow
        Case "pbTextAutoFitBestFit": PbTextAutoFitTypeFromString = pbTextAutoFitBestFit
        Case "pbTextAutoFitGrowToFit": PbTextAutoFitTypeFromString = pbTextAutoFitGrowToFit
    End Select
End Function

Function PbTextAutoFitTypeToString(value As PbTextAutoFitType) As String
    Select Case value
        Case pbTextAutoFitNone: PbTextAutoFitTypeToString = "pbTextAutoFitNone"
        Case pbTextAutoFitShrinkOnOverflow: PbTextAutoFitTypeToString = "pbTextAutoFitShrinkOnOverflow"
        Case pbTextAutoFitBestFit: PbTextAutoFitTypeToString = "pbTextAutoFitBestFit"
        Case pbTextAutoFitGrowToFit: PbTextAutoFitTypeToString = "pbTextAutoFitGrowToFit"
    End Select
End Function
