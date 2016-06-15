Attribute VB_Name = "wWdLanguageID2000"
Function WdLanguageID2000FromString(value As String) As WdLanguageID2000
    If IsNumeric(value) Then
        WdLanguageID2000FromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdChineseHongKong": WdLanguageID2000FromString = wdChineseHongKong
        Case "wdChineseMacao": WdLanguageID2000FromString = wdChineseMacao
        Case "wdEnglishTrinidad": WdLanguageID2000FromString = wdEnglishTrinidad
    End Select
End Function

Function WdLanguageID2000ToString(value As WdLanguageID2000) As String
    Select Case value
        Case wdChineseHongKong: WdLanguageID2000ToString = "wdChineseHongKong"
        Case wdChineseMacao: WdLanguageID2000ToString = "wdChineseMacao"
        Case wdEnglishTrinidad: WdLanguageID2000ToString = "wdEnglishTrinidad"
    End Select
End Function
