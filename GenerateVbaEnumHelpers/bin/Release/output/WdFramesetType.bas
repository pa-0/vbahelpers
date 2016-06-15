Attribute VB_Name = "wWdFramesetType"
Function WdFramesetTypeFromString(value As String) As WdFramesetType
    If IsNumeric(value) Then
        WdFramesetTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFramesetTypeFrameset": WdFramesetTypeFromString = wdFramesetTypeFrameset
        Case "wdFramesetTypeFrame": WdFramesetTypeFromString = wdFramesetTypeFrame
    End Select
End Function

Function WdFramesetTypeToString(value As WdFramesetType) As String
    Select Case value
        Case wdFramesetTypeFrameset: WdFramesetTypeToString = "wdFramesetTypeFrameset"
        Case wdFramesetTypeFrame: WdFramesetTypeToString = "wdFramesetTypeFrame"
    End Select
End Function
