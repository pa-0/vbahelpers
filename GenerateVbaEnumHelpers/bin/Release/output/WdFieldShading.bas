Attribute VB_Name = "wWdFieldShading"
Function WdFieldShadingFromString(value As String) As WdFieldShading
    If IsNumeric(value) Then
        WdFieldShadingFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFieldShadingNever": WdFieldShadingFromString = wdFieldShadingNever
        Case "wdFieldShadingAlways": WdFieldShadingFromString = wdFieldShadingAlways
        Case "wdFieldShadingWhenSelected": WdFieldShadingFromString = wdFieldShadingWhenSelected
    End Select
End Function

Function WdFieldShadingToString(value As WdFieldShading) As String
    Select Case value
        Case wdFieldShadingNever: WdFieldShadingToString = "wdFieldShadingNever"
        Case wdFieldShadingAlways: WdFieldShadingToString = "wdFieldShadingAlways"
        Case wdFieldShadingWhenSelected: WdFieldShadingToString = "wdFieldShadingWhenSelected"
    End Select
End Function
