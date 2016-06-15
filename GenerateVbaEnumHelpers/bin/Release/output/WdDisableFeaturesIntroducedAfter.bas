Attribute VB_Name = "wWdDisableFeaturesIntroducedAfter"
Function WdDisableFeaturesIntroducedAfterFromString(value As String) As WdDisableFeaturesIntroducedAfter
    If IsNumeric(value) Then
        WdDisableFeaturesIntroducedAfterFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wd70": WdDisableFeaturesIntroducedAfterFromString = wd70
        Case "wd70FE": WdDisableFeaturesIntroducedAfterFromString = wd70FE
        Case "wd80": WdDisableFeaturesIntroducedAfterFromString = wd80
    End Select
End Function

Function WdDisableFeaturesIntroducedAfterToString(value As WdDisableFeaturesIntroducedAfter) As String
    Select Case value
        Case wd70: WdDisableFeaturesIntroducedAfterToString = "wd70"
        Case wd70FE: WdDisableFeaturesIntroducedAfterToString = "wd70FE"
        Case wd80: WdDisableFeaturesIntroducedAfterToString = "wd80"
    End Select
End Function
