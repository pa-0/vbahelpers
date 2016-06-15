Attribute VB_Name = "wPbWrapType"
Function PbWrapTypeFromString(value As String) As PbWrapType
    If IsNumeric(value) Then
        PbWrapTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWrapTypeNone": PbWrapTypeFromString = pbWrapTypeNone
        Case "pbWrapTypeSquare": PbWrapTypeFromString = pbWrapTypeSquare
        Case "pbWrapTypeTight": PbWrapTypeFromString = pbWrapTypeTight
        Case "pbWrapTypeThrough": PbWrapTypeFromString = pbWrapTypeThrough
        Case "pbWrapTypeTopAndBottom": PbWrapTypeFromString = pbWrapTypeTopAndBottom
        Case "pbWrapTypeMixed": PbWrapTypeFromString = pbWrapTypeMixed
    End Select
End Function

Function PbWrapTypeToString(value As PbWrapType) As String
    Select Case value
        Case pbWrapTypeNone: PbWrapTypeToString = "pbWrapTypeNone"
        Case pbWrapTypeSquare: PbWrapTypeToString = "pbWrapTypeSquare"
        Case pbWrapTypeTight: PbWrapTypeToString = "pbWrapTypeTight"
        Case pbWrapTypeThrough: PbWrapTypeToString = "pbWrapTypeThrough"
        Case pbWrapTypeTopAndBottom: PbWrapTypeToString = "pbWrapTypeTopAndBottom"
        Case pbWrapTypeMixed: PbWrapTypeToString = "pbWrapTypeMixed"
    End Select
End Function
