Attribute VB_Name = "wMsoLanguageIDHidden"
Function MsoLanguageIDHiddenFromString(value As String) As MsoLanguageIDHidden
    If IsNumeric(value) Then
        MsoLanguageIDHiddenFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLanguageIDChineseHongKong": MsoLanguageIDHiddenFromString = msoLanguageIDChineseHongKong
        Case "msoLanguageIDChineseMacao": MsoLanguageIDHiddenFromString = msoLanguageIDChineseMacao
        Case "msoLanguageIDEnglishTrinidad": MsoLanguageIDHiddenFromString = msoLanguageIDEnglishTrinidad
    End Select
End Function

Function MsoLanguageIDHiddenToString(value As MsoLanguageIDHidden) As String
    Select Case value
        Case msoLanguageIDChineseHongKong: MsoLanguageIDHiddenToString = "msoLanguageIDChineseHongKong"
        Case msoLanguageIDChineseMacao: MsoLanguageIDHiddenToString = "msoLanguageIDChineseMacao"
        Case msoLanguageIDEnglishTrinidad: MsoLanguageIDHiddenToString = "msoLanguageIDEnglishTrinidad"
    End Select
End Function
