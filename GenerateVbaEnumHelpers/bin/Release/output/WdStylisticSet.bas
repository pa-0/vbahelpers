Attribute VB_Name = "wWdStylisticSet"
Function WdStylisticSetFromString(value As String) As WdStylisticSet
    If IsNumeric(value) Then
        WdStylisticSetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdStylisticSetDefault": WdStylisticSetFromString = wdStylisticSetDefault
        Case "wdStylisticSet01": WdStylisticSetFromString = wdStylisticSet01
        Case "wdStylisticSet02": WdStylisticSetFromString = wdStylisticSet02
        Case "wdStylisticSet03": WdStylisticSetFromString = wdStylisticSet03
        Case "wdStylisticSet04": WdStylisticSetFromString = wdStylisticSet04
        Case "wdStylisticSet05": WdStylisticSetFromString = wdStylisticSet05
        Case "wdStylisticSet06": WdStylisticSetFromString = wdStylisticSet06
        Case "wdStylisticSet07": WdStylisticSetFromString = wdStylisticSet07
        Case "wdStylisticSet08": WdStylisticSetFromString = wdStylisticSet08
        Case "wdStylisticSet09": WdStylisticSetFromString = wdStylisticSet09
        Case "wdStylisticSet10": WdStylisticSetFromString = wdStylisticSet10
        Case "wdStylisticSet11": WdStylisticSetFromString = wdStylisticSet11
        Case "wdStylisticSet12": WdStylisticSetFromString = wdStylisticSet12
        Case "wdStylisticSet13": WdStylisticSetFromString = wdStylisticSet13
        Case "wdStylisticSet14": WdStylisticSetFromString = wdStylisticSet14
        Case "wdStylisticSet15": WdStylisticSetFromString = wdStylisticSet15
        Case "wdStylisticSet16": WdStylisticSetFromString = wdStylisticSet16
        Case "wdStylisticSet17": WdStylisticSetFromString = wdStylisticSet17
        Case "wdStylisticSet18": WdStylisticSetFromString = wdStylisticSet18
        Case "wdStylisticSet19": WdStylisticSetFromString = wdStylisticSet19
        Case "wdStylisticSet20": WdStylisticSetFromString = wdStylisticSet20
    End Select
End Function

Function WdStylisticSetToString(value As WdStylisticSet) As String
    Select Case value
        Case wdStylisticSetDefault: WdStylisticSetToString = "wdStylisticSetDefault"
        Case wdStylisticSet01: WdStylisticSetToString = "wdStylisticSet01"
        Case wdStylisticSet02: WdStylisticSetToString = "wdStylisticSet02"
        Case wdStylisticSet03: WdStylisticSetToString = "wdStylisticSet03"
        Case wdStylisticSet04: WdStylisticSetToString = "wdStylisticSet04"
        Case wdStylisticSet05: WdStylisticSetToString = "wdStylisticSet05"
        Case wdStylisticSet06: WdStylisticSetToString = "wdStylisticSet06"
        Case wdStylisticSet07: WdStylisticSetToString = "wdStylisticSet07"
        Case wdStylisticSet08: WdStylisticSetToString = "wdStylisticSet08"
        Case wdStylisticSet09: WdStylisticSetToString = "wdStylisticSet09"
        Case wdStylisticSet10: WdStylisticSetToString = "wdStylisticSet10"
        Case wdStylisticSet11: WdStylisticSetToString = "wdStylisticSet11"
        Case wdStylisticSet12: WdStylisticSetToString = "wdStylisticSet12"
        Case wdStylisticSet13: WdStylisticSetToString = "wdStylisticSet13"
        Case wdStylisticSet14: WdStylisticSetToString = "wdStylisticSet14"
        Case wdStylisticSet15: WdStylisticSetToString = "wdStylisticSet15"
        Case wdStylisticSet16: WdStylisticSetToString = "wdStylisticSet16"
        Case wdStylisticSet17: WdStylisticSetToString = "wdStylisticSet17"
        Case wdStylisticSet18: WdStylisticSetToString = "wdStylisticSet18"
        Case wdStylisticSet19: WdStylisticSetToString = "wdStylisticSet19"
        Case wdStylisticSet20: WdStylisticSetToString = "wdStylisticSet20"
    End Select
End Function
