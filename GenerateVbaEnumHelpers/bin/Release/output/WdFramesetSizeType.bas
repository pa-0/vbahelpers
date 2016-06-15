Attribute VB_Name = "wWdFramesetSizeType"
Function WdFramesetSizeTypeFromString(value As String) As WdFramesetSizeType
    If IsNumeric(value) Then
        WdFramesetSizeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFramesetSizeTypePercent": WdFramesetSizeTypeFromString = wdFramesetSizeTypePercent
        Case "wdFramesetSizeTypeFixed": WdFramesetSizeTypeFromString = wdFramesetSizeTypeFixed
        Case "wdFramesetSizeTypeRelative": WdFramesetSizeTypeFromString = wdFramesetSizeTypeRelative
    End Select
End Function

Function WdFramesetSizeTypeToString(value As WdFramesetSizeType) As String
    Select Case value
        Case wdFramesetSizeTypePercent: WdFramesetSizeTypeToString = "wdFramesetSizeTypePercent"
        Case wdFramesetSizeTypeFixed: WdFramesetSizeTypeToString = "wdFramesetSizeTypeFixed"
        Case wdFramesetSizeTypeRelative: WdFramesetSizeTypeToString = "wdFramesetSizeTypeRelative"
    End Select
End Function
