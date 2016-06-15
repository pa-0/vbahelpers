Attribute VB_Name = "wWdLineSpacing"
Function WdLineSpacingFromString(value As String) As WdLineSpacing
    If IsNumeric(value) Then
        WdLineSpacingFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLineSpaceSingle": WdLineSpacingFromString = wdLineSpaceSingle
        Case "wdLineSpace1pt5": WdLineSpacingFromString = wdLineSpace1pt5
        Case "wdLineSpaceDouble": WdLineSpacingFromString = wdLineSpaceDouble
        Case "wdLineSpaceAtLeast": WdLineSpacingFromString = wdLineSpaceAtLeast
        Case "wdLineSpaceExactly": WdLineSpacingFromString = wdLineSpaceExactly
        Case "wdLineSpaceMultiple": WdLineSpacingFromString = wdLineSpaceMultiple
    End Select
End Function

Function WdLineSpacingToString(value As WdLineSpacing) As String
    Select Case value
        Case wdLineSpaceSingle: WdLineSpacingToString = "wdLineSpaceSingle"
        Case wdLineSpace1pt5: WdLineSpacingToString = "wdLineSpace1pt5"
        Case wdLineSpaceDouble: WdLineSpacingToString = "wdLineSpaceDouble"
        Case wdLineSpaceAtLeast: WdLineSpacingToString = "wdLineSpaceAtLeast"
        Case wdLineSpaceExactly: WdLineSpacingToString = "wdLineSpaceExactly"
        Case wdLineSpaceMultiple: WdLineSpacingToString = "wdLineSpaceMultiple"
    End Select
End Function
