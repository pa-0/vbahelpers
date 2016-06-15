Attribute VB_Name = "wPbSchemeColorIndex"
Function PbSchemeColorIndexFromString(value As String) As PbSchemeColorIndex
    If IsNumeric(value) Then
        PbSchemeColorIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbSchemeColorNone": PbSchemeColorIndexFromString = pbSchemeColorNone
        Case "pbSchemeColorMain": PbSchemeColorIndexFromString = pbSchemeColorMain
        Case "pbSchemeColorAccent1": PbSchemeColorIndexFromString = pbSchemeColorAccent1
        Case "pbSchemeColorAccent2": PbSchemeColorIndexFromString = pbSchemeColorAccent2
        Case "pbSchemeColorAccent3": PbSchemeColorIndexFromString = pbSchemeColorAccent3
        Case "pbSchemeColorAccent4": PbSchemeColorIndexFromString = pbSchemeColorAccent4
        Case "pbSchemeColorHyperlink": PbSchemeColorIndexFromString = pbSchemeColorHyperlink
        Case "pbSchemeColorFollowedHyperlink": PbSchemeColorIndexFromString = pbSchemeColorFollowedHyperlink
        Case "pbSchemeColorAccent5": PbSchemeColorIndexFromString = pbSchemeColorAccent5
    End Select
End Function

Function PbSchemeColorIndexToString(value As PbSchemeColorIndex) As String
    Select Case value
        Case pbSchemeColorNone: PbSchemeColorIndexToString = "pbSchemeColorNone"
        Case pbSchemeColorMain: PbSchemeColorIndexToString = "pbSchemeColorMain"
        Case pbSchemeColorAccent1: PbSchemeColorIndexToString = "pbSchemeColorAccent1"
        Case pbSchemeColorAccent2: PbSchemeColorIndexToString = "pbSchemeColorAccent2"
        Case pbSchemeColorAccent3: PbSchemeColorIndexToString = "pbSchemeColorAccent3"
        Case pbSchemeColorAccent4: PbSchemeColorIndexToString = "pbSchemeColorAccent4"
        Case pbSchemeColorHyperlink: PbSchemeColorIndexToString = "pbSchemeColorHyperlink"
        Case pbSchemeColorFollowedHyperlink: PbSchemeColorIndexToString = "pbSchemeColorFollowedHyperlink"
        Case pbSchemeColorAccent5: PbSchemeColorIndexToString = "pbSchemeColorAccent5"
    End Select
End Function
