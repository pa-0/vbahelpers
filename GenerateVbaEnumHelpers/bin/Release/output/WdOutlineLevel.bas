Attribute VB_Name = "wWdOutlineLevel"
Function WdOutlineLevelFromString(value As String) As WdOutlineLevel
    If IsNumeric(value) Then
        WdOutlineLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOutlineLevel1": WdOutlineLevelFromString = wdOutlineLevel1
        Case "wdOutlineLevel2": WdOutlineLevelFromString = wdOutlineLevel2
        Case "wdOutlineLevel3": WdOutlineLevelFromString = wdOutlineLevel3
        Case "wdOutlineLevel4": WdOutlineLevelFromString = wdOutlineLevel4
        Case "wdOutlineLevel5": WdOutlineLevelFromString = wdOutlineLevel5
        Case "wdOutlineLevel6": WdOutlineLevelFromString = wdOutlineLevel6
        Case "wdOutlineLevel7": WdOutlineLevelFromString = wdOutlineLevel7
        Case "wdOutlineLevel8": WdOutlineLevelFromString = wdOutlineLevel8
        Case "wdOutlineLevel9": WdOutlineLevelFromString = wdOutlineLevel9
        Case "wdOutlineLevelBodyText": WdOutlineLevelFromString = wdOutlineLevelBodyText
    End Select
End Function

Function WdOutlineLevelToString(value As WdOutlineLevel) As String
    Select Case value
        Case wdOutlineLevel1: WdOutlineLevelToString = "wdOutlineLevel1"
        Case wdOutlineLevel2: WdOutlineLevelToString = "wdOutlineLevel2"
        Case wdOutlineLevel3: WdOutlineLevelToString = "wdOutlineLevel3"
        Case wdOutlineLevel4: WdOutlineLevelToString = "wdOutlineLevel4"
        Case wdOutlineLevel5: WdOutlineLevelToString = "wdOutlineLevel5"
        Case wdOutlineLevel6: WdOutlineLevelToString = "wdOutlineLevel6"
        Case wdOutlineLevel7: WdOutlineLevelToString = "wdOutlineLevel7"
        Case wdOutlineLevel8: WdOutlineLevelToString = "wdOutlineLevel8"
        Case wdOutlineLevel9: WdOutlineLevelToString = "wdOutlineLevel9"
        Case wdOutlineLevelBodyText: WdOutlineLevelToString = "wdOutlineLevelBodyText"
    End Select
End Function
