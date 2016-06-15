Attribute VB_Name = "wXlDataLabelPosition"
Function XlDataLabelPositionFromString(value As String) As XlDataLabelPosition
    If IsNumeric(value) Then
        XlDataLabelPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlLabelPositionAbove": XlDataLabelPositionFromString = xlLabelPositionAbove
        Case "xlLabelPositionBelow": XlDataLabelPositionFromString = xlLabelPositionBelow
        Case "xlLabelPositionOutsideEnd": XlDataLabelPositionFromString = xlLabelPositionOutsideEnd
        Case "xlLabelPositionInsideEnd": XlDataLabelPositionFromString = xlLabelPositionInsideEnd
        Case "xlLabelPositionInsideBase": XlDataLabelPositionFromString = xlLabelPositionInsideBase
        Case "xlLabelPositionBestFit": XlDataLabelPositionFromString = xlLabelPositionBestFit
        Case "xlLabelPositionMixed": XlDataLabelPositionFromString = xlLabelPositionMixed
        Case "xlLabelPositionCustom": XlDataLabelPositionFromString = xlLabelPositionCustom
        Case "xlLabelPositionRight": XlDataLabelPositionFromString = xlLabelPositionRight
        Case "xlLabelPositionLeft": XlDataLabelPositionFromString = xlLabelPositionLeft
        Case "xlLabelPositionCenter": XlDataLabelPositionFromString = xlLabelPositionCenter
    End Select
End Function

Function XlDataLabelPositionToString(value As XlDataLabelPosition) As String
    Select Case value
        Case xlLabelPositionAbove: XlDataLabelPositionToString = "xlLabelPositionAbove"
        Case xlLabelPositionBelow: XlDataLabelPositionToString = "xlLabelPositionBelow"
        Case xlLabelPositionOutsideEnd: XlDataLabelPositionToString = "xlLabelPositionOutsideEnd"
        Case xlLabelPositionInsideEnd: XlDataLabelPositionToString = "xlLabelPositionInsideEnd"
        Case xlLabelPositionInsideBase: XlDataLabelPositionToString = "xlLabelPositionInsideBase"
        Case xlLabelPositionBestFit: XlDataLabelPositionToString = "xlLabelPositionBestFit"
        Case xlLabelPositionMixed: XlDataLabelPositionToString = "xlLabelPositionMixed"
        Case xlLabelPositionCustom: XlDataLabelPositionToString = "xlLabelPositionCustom"
        Case xlLabelPositionRight: XlDataLabelPositionToString = "xlLabelPositionRight"
        Case xlLabelPositionLeft: XlDataLabelPositionToString = "xlLabelPositionLeft"
        Case xlLabelPositionCenter: XlDataLabelPositionToString = "xlLabelPositionCenter"
    End Select
End Function
