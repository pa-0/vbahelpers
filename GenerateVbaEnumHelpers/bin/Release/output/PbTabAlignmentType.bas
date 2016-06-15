Attribute VB_Name = "wPbTabAlignmentType"
Function PbTabAlignmentTypeFromString(value As String) As PbTabAlignmentType
    If IsNumeric(value) Then
        PbTabAlignmentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTabAlignmentLeading": PbTabAlignmentTypeFromString = pbTabAlignmentLeading
        Case "pbTabAlignmentCenter": PbTabAlignmentTypeFromString = pbTabAlignmentCenter
        Case "pbTabAlignmentTrailing": PbTabAlignmentTypeFromString = pbTabAlignmentTrailing
        Case "pbTabAlignmentDecimal": PbTabAlignmentTypeFromString = pbTabAlignmentDecimal
    End Select
End Function

Function PbTabAlignmentTypeToString(value As PbTabAlignmentType) As String
    Select Case value
        Case pbTabAlignmentLeading: PbTabAlignmentTypeToString = "pbTabAlignmentLeading"
        Case pbTabAlignmentCenter: PbTabAlignmentTypeToString = "pbTabAlignmentCenter"
        Case pbTabAlignmentTrailing: PbTabAlignmentTypeToString = "pbTabAlignmentTrailing"
        Case pbTabAlignmentDecimal: PbTabAlignmentTypeToString = "pbTabAlignmentDecimal"
    End Select
End Function
