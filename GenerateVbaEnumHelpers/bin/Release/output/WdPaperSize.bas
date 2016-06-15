Attribute VB_Name = "wWdPaperSize"
Function WdPaperSizeFromString(value As String) As WdPaperSize
    If IsNumeric(value) Then
        WdPaperSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPaper10x14": WdPaperSizeFromString = wdPaper10x14
        Case "wdPaper11x17": WdPaperSizeFromString = wdPaper11x17
        Case "wdPaperLetter": WdPaperSizeFromString = wdPaperLetter
        Case "wdPaperLetterSmall": WdPaperSizeFromString = wdPaperLetterSmall
        Case "wdPaperLegal": WdPaperSizeFromString = wdPaperLegal
        Case "wdPaperExecutive": WdPaperSizeFromString = wdPaperExecutive
        Case "wdPaperA3": WdPaperSizeFromString = wdPaperA3
        Case "wdPaperA4": WdPaperSizeFromString = wdPaperA4
        Case "wdPaperA4Small": WdPaperSizeFromString = wdPaperA4Small
        Case "wdPaperA5": WdPaperSizeFromString = wdPaperA5
        Case "wdPaperB4": WdPaperSizeFromString = wdPaperB4
        Case "wdPaperB5": WdPaperSizeFromString = wdPaperB5
        Case "wdPaperCSheet": WdPaperSizeFromString = wdPaperCSheet
        Case "wdPaperDSheet": WdPaperSizeFromString = wdPaperDSheet
        Case "wdPaperESheet": WdPaperSizeFromString = wdPaperESheet
        Case "wdPaperFanfoldLegalGerman": WdPaperSizeFromString = wdPaperFanfoldLegalGerman
        Case "wdPaperFanfoldStdGerman": WdPaperSizeFromString = wdPaperFanfoldStdGerman
        Case "wdPaperFanfoldUS": WdPaperSizeFromString = wdPaperFanfoldUS
        Case "wdPaperFolio": WdPaperSizeFromString = wdPaperFolio
        Case "wdPaperLedger": WdPaperSizeFromString = wdPaperLedger
        Case "wdPaperNote": WdPaperSizeFromString = wdPaperNote
        Case "wdPaperQuarto": WdPaperSizeFromString = wdPaperQuarto
        Case "wdPaperStatement": WdPaperSizeFromString = wdPaperStatement
        Case "wdPaperTabloid": WdPaperSizeFromString = wdPaperTabloid
        Case "wdPaperEnvelope9": WdPaperSizeFromString = wdPaperEnvelope9
        Case "wdPaperEnvelope10": WdPaperSizeFromString = wdPaperEnvelope10
        Case "wdPaperEnvelope11": WdPaperSizeFromString = wdPaperEnvelope11
        Case "wdPaperEnvelope12": WdPaperSizeFromString = wdPaperEnvelope12
        Case "wdPaperEnvelope14": WdPaperSizeFromString = wdPaperEnvelope14
        Case "wdPaperEnvelopeB4": WdPaperSizeFromString = wdPaperEnvelopeB4
        Case "wdPaperEnvelopeB5": WdPaperSizeFromString = wdPaperEnvelopeB5
        Case "wdPaperEnvelopeB6": WdPaperSizeFromString = wdPaperEnvelopeB6
        Case "wdPaperEnvelopeC3": WdPaperSizeFromString = wdPaperEnvelopeC3
        Case "wdPaperEnvelopeC4": WdPaperSizeFromString = wdPaperEnvelopeC4
        Case "wdPaperEnvelopeC5": WdPaperSizeFromString = wdPaperEnvelopeC5
        Case "wdPaperEnvelopeC6": WdPaperSizeFromString = wdPaperEnvelopeC6
        Case "wdPaperEnvelopeC65": WdPaperSizeFromString = wdPaperEnvelopeC65
        Case "wdPaperEnvelopeDL": WdPaperSizeFromString = wdPaperEnvelopeDL
        Case "wdPaperEnvelopeItaly": WdPaperSizeFromString = wdPaperEnvelopeItaly
        Case "wdPaperEnvelopeMonarch": WdPaperSizeFromString = wdPaperEnvelopeMonarch
        Case "wdPaperEnvelopePersonal": WdPaperSizeFromString = wdPaperEnvelopePersonal
        Case "wdPaperCustom": WdPaperSizeFromString = wdPaperCustom
    End Select
End Function

Function WdPaperSizeToString(value As WdPaperSize) As String
    Select Case value
        Case wdPaper10x14: WdPaperSizeToString = "wdPaper10x14"
        Case wdPaper11x17: WdPaperSizeToString = "wdPaper11x17"
        Case wdPaperLetter: WdPaperSizeToString = "wdPaperLetter"
        Case wdPaperLetterSmall: WdPaperSizeToString = "wdPaperLetterSmall"
        Case wdPaperLegal: WdPaperSizeToString = "wdPaperLegal"
        Case wdPaperExecutive: WdPaperSizeToString = "wdPaperExecutive"
        Case wdPaperA3: WdPaperSizeToString = "wdPaperA3"
        Case wdPaperA4: WdPaperSizeToString = "wdPaperA4"
        Case wdPaperA4Small: WdPaperSizeToString = "wdPaperA4Small"
        Case wdPaperA5: WdPaperSizeToString = "wdPaperA5"
        Case wdPaperB4: WdPaperSizeToString = "wdPaperB4"
        Case wdPaperB5: WdPaperSizeToString = "wdPaperB5"
        Case wdPaperCSheet: WdPaperSizeToString = "wdPaperCSheet"
        Case wdPaperDSheet: WdPaperSizeToString = "wdPaperDSheet"
        Case wdPaperESheet: WdPaperSizeToString = "wdPaperESheet"
        Case wdPaperFanfoldLegalGerman: WdPaperSizeToString = "wdPaperFanfoldLegalGerman"
        Case wdPaperFanfoldStdGerman: WdPaperSizeToString = "wdPaperFanfoldStdGerman"
        Case wdPaperFanfoldUS: WdPaperSizeToString = "wdPaperFanfoldUS"
        Case wdPaperFolio: WdPaperSizeToString = "wdPaperFolio"
        Case wdPaperLedger: WdPaperSizeToString = "wdPaperLedger"
        Case wdPaperNote: WdPaperSizeToString = "wdPaperNote"
        Case wdPaperQuarto: WdPaperSizeToString = "wdPaperQuarto"
        Case wdPaperStatement: WdPaperSizeToString = "wdPaperStatement"
        Case wdPaperTabloid: WdPaperSizeToString = "wdPaperTabloid"
        Case wdPaperEnvelope9: WdPaperSizeToString = "wdPaperEnvelope9"
        Case wdPaperEnvelope10: WdPaperSizeToString = "wdPaperEnvelope10"
        Case wdPaperEnvelope11: WdPaperSizeToString = "wdPaperEnvelope11"
        Case wdPaperEnvelope12: WdPaperSizeToString = "wdPaperEnvelope12"
        Case wdPaperEnvelope14: WdPaperSizeToString = "wdPaperEnvelope14"
        Case wdPaperEnvelopeB4: WdPaperSizeToString = "wdPaperEnvelopeB4"
        Case wdPaperEnvelopeB5: WdPaperSizeToString = "wdPaperEnvelopeB5"
        Case wdPaperEnvelopeB6: WdPaperSizeToString = "wdPaperEnvelopeB6"
        Case wdPaperEnvelopeC3: WdPaperSizeToString = "wdPaperEnvelopeC3"
        Case wdPaperEnvelopeC4: WdPaperSizeToString = "wdPaperEnvelopeC4"
        Case wdPaperEnvelopeC5: WdPaperSizeToString = "wdPaperEnvelopeC5"
        Case wdPaperEnvelopeC6: WdPaperSizeToString = "wdPaperEnvelopeC6"
        Case wdPaperEnvelopeC65: WdPaperSizeToString = "wdPaperEnvelopeC65"
        Case wdPaperEnvelopeDL: WdPaperSizeToString = "wdPaperEnvelopeDL"
        Case wdPaperEnvelopeItaly: WdPaperSizeToString = "wdPaperEnvelopeItaly"
        Case wdPaperEnvelopeMonarch: WdPaperSizeToString = "wdPaperEnvelopeMonarch"
        Case wdPaperEnvelopePersonal: WdPaperSizeToString = "wdPaperEnvelopePersonal"
        Case wdPaperCustom: WdPaperSizeToString = "wdPaperCustom"
    End Select
End Function
