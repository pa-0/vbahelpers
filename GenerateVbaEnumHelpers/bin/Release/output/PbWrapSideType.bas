Attribute VB_Name = "wPbWrapSideType"
Function PbWrapSideTypeFromString(value As String) As PbWrapSideType
    If IsNumeric(value) Then
        PbWrapSideTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWrapSideBoth": PbWrapSideTypeFromString = pbWrapSideBoth
        Case "pbWrapSideLeft": PbWrapSideTypeFromString = pbWrapSideLeft
        Case "pbWrapSideRight": PbWrapSideTypeFromString = pbWrapSideRight
        Case "pbWrapSideLarger": PbWrapSideTypeFromString = pbWrapSideLarger
        Case "pbWrapSideNeither": PbWrapSideTypeFromString = pbWrapSideNeither
        Case "pbWrapSideMixed": PbWrapSideTypeFromString = pbWrapSideMixed
    End Select
End Function

Function PbWrapSideTypeToString(value As PbWrapSideType) As String
    Select Case value
        Case pbWrapSideBoth: PbWrapSideTypeToString = "pbWrapSideBoth"
        Case pbWrapSideLeft: PbWrapSideTypeToString = "pbWrapSideLeft"
        Case pbWrapSideRight: PbWrapSideTypeToString = "pbWrapSideRight"
        Case pbWrapSideLarger: PbWrapSideTypeToString = "pbWrapSideLarger"
        Case pbWrapSideNeither: PbWrapSideTypeToString = "pbWrapSideNeither"
        Case pbWrapSideMixed: PbWrapSideTypeToString = "pbWrapSideMixed"
    End Select
End Function
