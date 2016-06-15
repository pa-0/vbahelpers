Attribute VB_Name = "wPbFontScriptType"
Function PbFontScriptTypeFromString(value As String) As PbFontScriptType
    If IsNumeric(value) Then
        PbFontScriptTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFontScriptDefault": PbFontScriptTypeFromString = pbFontScriptDefault
        Case "pbFontScriptAsciiLatin": PbFontScriptTypeFromString = pbFontScriptAsciiLatin
        Case "pbFontScriptLatin": PbFontScriptTypeFromString = pbFontScriptLatin
        Case "pbFontScriptGreek": PbFontScriptTypeFromString = pbFontScriptGreek
        Case "pbFontScriptCyrillic": PbFontScriptTypeFromString = pbFontScriptCyrillic
        Case "pbFontScriptArmenian": PbFontScriptTypeFromString = pbFontScriptArmenian
        Case "pbFontScriptHebrew": PbFontScriptTypeFromString = pbFontScriptHebrew
        Case "pbFontScriptArabic": PbFontScriptTypeFromString = pbFontScriptArabic
        Case "pbFontScriptDevanagari": PbFontScriptTypeFromString = pbFontScriptDevanagari
        Case "pbFontScriptBengali": PbFontScriptTypeFromString = pbFontScriptBengali
        Case "pbFontScriptGurmukhi": PbFontScriptTypeFromString = pbFontScriptGurmukhi
        Case "pbFontScriptGujarati": PbFontScriptTypeFromString = pbFontScriptGujarati
        Case "pbFontScriptOriya": PbFontScriptTypeFromString = pbFontScriptOriya
        Case "pbFontScriptTamil": PbFontScriptTypeFromString = pbFontScriptTamil
        Case "pbFontScriptTelugu": PbFontScriptTypeFromString = pbFontScriptTelugu
        Case "pbFontScriptKannada": PbFontScriptTypeFromString = pbFontScriptKannada
        Case "pbFontScriptMalayalam": PbFontScriptTypeFromString = pbFontScriptMalayalam
        Case "pbFontScriptThai": PbFontScriptTypeFromString = pbFontScriptThai
        Case "pbFontScriptLao": PbFontScriptTypeFromString = pbFontScriptLao
        Case "pbFontScriptTibetan": PbFontScriptTypeFromString = pbFontScriptTibetan
        Case "pbFontScriptGeorgian": PbFontScriptTypeFromString = pbFontScriptGeorgian
        Case "pbFontScriptHangul": PbFontScriptTypeFromString = pbFontScriptHangul
        Case "pbFontScriptKana": PbFontScriptTypeFromString = pbFontScriptKana
        Case "pbFontScriptBopomofo": PbFontScriptTypeFromString = pbFontScriptBopomofo
        Case "pbFontScriptHan": PbFontScriptTypeFromString = pbFontScriptHan
        Case "pbFontScriptHalfWidthKana": PbFontScriptTypeFromString = pbFontScriptHalfWidthKana
        Case "pbFontScriptEUDC": PbFontScriptTypeFromString = pbFontScriptEUDC
        Case "pbFontScriptYi": PbFontScriptTypeFromString = pbFontScriptYi
        Case "pbFontScriptHanSurrogate": PbFontScriptTypeFromString = pbFontScriptHanSurrogate
        Case "pbFontScriptNonHanSurrogate": PbFontScriptTypeFromString = pbFontScriptNonHanSurrogate
        Case "pbFontScriptSyriac": PbFontScriptTypeFromString = pbFontScriptSyriac
        Case "pbFontScriptThaana": PbFontScriptTypeFromString = pbFontScriptThaana
        Case "pbFontScriptMyanmar": PbFontScriptTypeFromString = pbFontScriptMyanmar
        Case "pbFontScriptSinhala": PbFontScriptTypeFromString = pbFontScriptSinhala
        Case "pbFontScriptEthiopic": PbFontScriptTypeFromString = pbFontScriptEthiopic
        Case "pbFontScriptCherokee": PbFontScriptTypeFromString = pbFontScriptCherokee
        Case "pbFontScriptCanadianAbor": PbFontScriptTypeFromString = pbFontScriptCanadianAbor
        Case "pbFontScriptOgham": PbFontScriptTypeFromString = pbFontScriptOgham
        Case "pbFontScriptRunic": PbFontScriptTypeFromString = pbFontScriptRunic
        Case "pbFontScriptKhmer": PbFontScriptTypeFromString = pbFontScriptKhmer
        Case "pbFontScriptMongolian": PbFontScriptTypeFromString = pbFontScriptMongolian
        Case "pbFontScriptBraille": PbFontScriptTypeFromString = pbFontScriptBraille
        Case "pbFontScriptCurrency": PbFontScriptTypeFromString = pbFontScriptCurrency
        Case "pbFontScriptAsciiSym": PbFontScriptTypeFromString = pbFontScriptAsciiSym
        Case "pbFontScriptMixed": PbFontScriptTypeFromString = pbFontScriptMixed
    End Select
End Function

Function PbFontScriptTypeToString(value As PbFontScriptType) As String
    Select Case value
        Case pbFontScriptDefault: PbFontScriptTypeToString = "pbFontScriptDefault"
        Case pbFontScriptAsciiLatin: PbFontScriptTypeToString = "pbFontScriptAsciiLatin"
        Case pbFontScriptLatin: PbFontScriptTypeToString = "pbFontScriptLatin"
        Case pbFontScriptGreek: PbFontScriptTypeToString = "pbFontScriptGreek"
        Case pbFontScriptCyrillic: PbFontScriptTypeToString = "pbFontScriptCyrillic"
        Case pbFontScriptArmenian: PbFontScriptTypeToString = "pbFontScriptArmenian"
        Case pbFontScriptHebrew: PbFontScriptTypeToString = "pbFontScriptHebrew"
        Case pbFontScriptArabic: PbFontScriptTypeToString = "pbFontScriptArabic"
        Case pbFontScriptDevanagari: PbFontScriptTypeToString = "pbFontScriptDevanagari"
        Case pbFontScriptBengali: PbFontScriptTypeToString = "pbFontScriptBengali"
        Case pbFontScriptGurmukhi: PbFontScriptTypeToString = "pbFontScriptGurmukhi"
        Case pbFontScriptGujarati: PbFontScriptTypeToString = "pbFontScriptGujarati"
        Case pbFontScriptOriya: PbFontScriptTypeToString = "pbFontScriptOriya"
        Case pbFontScriptTamil: PbFontScriptTypeToString = "pbFontScriptTamil"
        Case pbFontScriptTelugu: PbFontScriptTypeToString = "pbFontScriptTelugu"
        Case pbFontScriptKannada: PbFontScriptTypeToString = "pbFontScriptKannada"
        Case pbFontScriptMalayalam: PbFontScriptTypeToString = "pbFontScriptMalayalam"
        Case pbFontScriptThai: PbFontScriptTypeToString = "pbFontScriptThai"
        Case pbFontScriptLao: PbFontScriptTypeToString = "pbFontScriptLao"
        Case pbFontScriptTibetan: PbFontScriptTypeToString = "pbFontScriptTibetan"
        Case pbFontScriptGeorgian: PbFontScriptTypeToString = "pbFontScriptGeorgian"
        Case pbFontScriptHangul: PbFontScriptTypeToString = "pbFontScriptHangul"
        Case pbFontScriptKana: PbFontScriptTypeToString = "pbFontScriptKana"
        Case pbFontScriptBopomofo: PbFontScriptTypeToString = "pbFontScriptBopomofo"
        Case pbFontScriptHan: PbFontScriptTypeToString = "pbFontScriptHan"
        Case pbFontScriptHalfWidthKana: PbFontScriptTypeToString = "pbFontScriptHalfWidthKana"
        Case pbFontScriptEUDC: PbFontScriptTypeToString = "pbFontScriptEUDC"
        Case pbFontScriptYi: PbFontScriptTypeToString = "pbFontScriptYi"
        Case pbFontScriptHanSurrogate: PbFontScriptTypeToString = "pbFontScriptHanSurrogate"
        Case pbFontScriptNonHanSurrogate: PbFontScriptTypeToString = "pbFontScriptNonHanSurrogate"
        Case pbFontScriptSyriac: PbFontScriptTypeToString = "pbFontScriptSyriac"
        Case pbFontScriptThaana: PbFontScriptTypeToString = "pbFontScriptThaana"
        Case pbFontScriptMyanmar: PbFontScriptTypeToString = "pbFontScriptMyanmar"
        Case pbFontScriptSinhala: PbFontScriptTypeToString = "pbFontScriptSinhala"
        Case pbFontScriptEthiopic: PbFontScriptTypeToString = "pbFontScriptEthiopic"
        Case pbFontScriptCherokee: PbFontScriptTypeToString = "pbFontScriptCherokee"
        Case pbFontScriptCanadianAbor: PbFontScriptTypeToString = "pbFontScriptCanadianAbor"
        Case pbFontScriptOgham: PbFontScriptTypeToString = "pbFontScriptOgham"
        Case pbFontScriptRunic: PbFontScriptTypeToString = "pbFontScriptRunic"
        Case pbFontScriptKhmer: PbFontScriptTypeToString = "pbFontScriptKhmer"
        Case pbFontScriptMongolian: PbFontScriptTypeToString = "pbFontScriptMongolian"
        Case pbFontScriptBraille: PbFontScriptTypeToString = "pbFontScriptBraille"
        Case pbFontScriptCurrency: PbFontScriptTypeToString = "pbFontScriptCurrency"
        Case pbFontScriptAsciiSym: PbFontScriptTypeToString = "pbFontScriptAsciiSym"
        Case pbFontScriptMixed: PbFontScriptTypeToString = "pbFontScriptMixed"
    End Select
End Function
