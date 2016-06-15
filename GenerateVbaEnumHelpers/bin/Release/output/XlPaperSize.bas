Attribute VB_Name = "wXlPaperSize"
Function XlPaperSizeFromString(value As String) As XlPaperSize
    If IsNumeric(value) Then
        XlPaperSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPaperLetter": XlPaperSizeFromString = xlPaperLetter
        Case "xlPaperLetterSmall": XlPaperSizeFromString = xlPaperLetterSmall
        Case "xlPaperTabloid": XlPaperSizeFromString = xlPaperTabloid
        Case "xlPaperLedger": XlPaperSizeFromString = xlPaperLedger
        Case "xlPaperLegal": XlPaperSizeFromString = xlPaperLegal
        Case "xlPaperStatement": XlPaperSizeFromString = xlPaperStatement
        Case "xlPaperExecutive": XlPaperSizeFromString = xlPaperExecutive
        Case "xlPaperA3": XlPaperSizeFromString = xlPaperA3
        Case "xlPaperA4": XlPaperSizeFromString = xlPaperA4
        Case "xlPaperA4Small": XlPaperSizeFromString = xlPaperA4Small
        Case "xlPaperA5": XlPaperSizeFromString = xlPaperA5
        Case "xlPaperB4": XlPaperSizeFromString = xlPaperB4
        Case "xlPaperB5": XlPaperSizeFromString = xlPaperB5
        Case "xlPaperFolio": XlPaperSizeFromString = xlPaperFolio
        Case "xlPaperQuarto": XlPaperSizeFromString = xlPaperQuarto
        Case "xlPaper10x14": XlPaperSizeFromString = xlPaper10x14
        Case "xlPaper11x17": XlPaperSizeFromString = xlPaper11x17
        Case "xlPaperNote": XlPaperSizeFromString = xlPaperNote
        Case "xlPaperEnvelope9": XlPaperSizeFromString = xlPaperEnvelope9
        Case "xlPaperEnvelope10": XlPaperSizeFromString = xlPaperEnvelope10
        Case "xlPaperEnvelope11": XlPaperSizeFromString = xlPaperEnvelope11
        Case "xlPaperEnvelope12": XlPaperSizeFromString = xlPaperEnvelope12
        Case "xlPaperEnvelope14": XlPaperSizeFromString = xlPaperEnvelope14
        Case "xlPaperCsheet": XlPaperSizeFromString = xlPaperCsheet
        Case "xlPaperDsheet": XlPaperSizeFromString = xlPaperDsheet
        Case "xlPaperEsheet": XlPaperSizeFromString = xlPaperEsheet
        Case "xlPaperEnvelopeDL": XlPaperSizeFromString = xlPaperEnvelopeDL
        Case "xlPaperEnvelopeC5": XlPaperSizeFromString = xlPaperEnvelopeC5
        Case "xlPaperEnvelopeC3": XlPaperSizeFromString = xlPaperEnvelopeC3
        Case "xlPaperEnvelopeC4": XlPaperSizeFromString = xlPaperEnvelopeC4
        Case "xlPaperEnvelopeC6": XlPaperSizeFromString = xlPaperEnvelopeC6
        Case "xlPaperEnvelopeC65": XlPaperSizeFromString = xlPaperEnvelopeC65
        Case "xlPaperEnvelopeB4": XlPaperSizeFromString = xlPaperEnvelopeB4
        Case "xlPaperEnvelopeB5": XlPaperSizeFromString = xlPaperEnvelopeB5
        Case "xlPaperEnvelopeB6": XlPaperSizeFromString = xlPaperEnvelopeB6
        Case "xlPaperEnvelopeItaly": XlPaperSizeFromString = xlPaperEnvelopeItaly
        Case "xlPaperEnvelopeMonarch": XlPaperSizeFromString = xlPaperEnvelopeMonarch
        Case "xlPaperEnvelopePersonal": XlPaperSizeFromString = xlPaperEnvelopePersonal
        Case "xlPaperFanfoldUS": XlPaperSizeFromString = xlPaperFanfoldUS
        Case "xlPaperFanfoldStdGerman": XlPaperSizeFromString = xlPaperFanfoldStdGerman
        Case "xlPaperFanfoldLegalGerman": XlPaperSizeFromString = xlPaperFanfoldLegalGerman
        Case "xlPaperUser": XlPaperSizeFromString = xlPaperUser
    End Select
End Function

Function XlPaperSizeToString(value As XlPaperSize) As String
    Select Case value
        Case xlPaperLetter: XlPaperSizeToString = "xlPaperLetter"
        Case xlPaperLetterSmall: XlPaperSizeToString = "xlPaperLetterSmall"
        Case xlPaperTabloid: XlPaperSizeToString = "xlPaperTabloid"
        Case xlPaperLedger: XlPaperSizeToString = "xlPaperLedger"
        Case xlPaperLegal: XlPaperSizeToString = "xlPaperLegal"
        Case xlPaperStatement: XlPaperSizeToString = "xlPaperStatement"
        Case xlPaperExecutive: XlPaperSizeToString = "xlPaperExecutive"
        Case xlPaperA3: XlPaperSizeToString = "xlPaperA3"
        Case xlPaperA4: XlPaperSizeToString = "xlPaperA4"
        Case xlPaperA4Small: XlPaperSizeToString = "xlPaperA4Small"
        Case xlPaperA5: XlPaperSizeToString = "xlPaperA5"
        Case xlPaperB4: XlPaperSizeToString = "xlPaperB4"
        Case xlPaperB5: XlPaperSizeToString = "xlPaperB5"
        Case xlPaperFolio: XlPaperSizeToString = "xlPaperFolio"
        Case xlPaperQuarto: XlPaperSizeToString = "xlPaperQuarto"
        Case xlPaper10x14: XlPaperSizeToString = "xlPaper10x14"
        Case xlPaper11x17: XlPaperSizeToString = "xlPaper11x17"
        Case xlPaperNote: XlPaperSizeToString = "xlPaperNote"
        Case xlPaperEnvelope9: XlPaperSizeToString = "xlPaperEnvelope9"
        Case xlPaperEnvelope10: XlPaperSizeToString = "xlPaperEnvelope10"
        Case xlPaperEnvelope11: XlPaperSizeToString = "xlPaperEnvelope11"
        Case xlPaperEnvelope12: XlPaperSizeToString = "xlPaperEnvelope12"
        Case xlPaperEnvelope14: XlPaperSizeToString = "xlPaperEnvelope14"
        Case xlPaperCsheet: XlPaperSizeToString = "xlPaperCsheet"
        Case xlPaperDsheet: XlPaperSizeToString = "xlPaperDsheet"
        Case xlPaperEsheet: XlPaperSizeToString = "xlPaperEsheet"
        Case xlPaperEnvelopeDL: XlPaperSizeToString = "xlPaperEnvelopeDL"
        Case xlPaperEnvelopeC5: XlPaperSizeToString = "xlPaperEnvelopeC5"
        Case xlPaperEnvelopeC3: XlPaperSizeToString = "xlPaperEnvelopeC3"
        Case xlPaperEnvelopeC4: XlPaperSizeToString = "xlPaperEnvelopeC4"
        Case xlPaperEnvelopeC6: XlPaperSizeToString = "xlPaperEnvelopeC6"
        Case xlPaperEnvelopeC65: XlPaperSizeToString = "xlPaperEnvelopeC65"
        Case xlPaperEnvelopeB4: XlPaperSizeToString = "xlPaperEnvelopeB4"
        Case xlPaperEnvelopeB5: XlPaperSizeToString = "xlPaperEnvelopeB5"
        Case xlPaperEnvelopeB6: XlPaperSizeToString = "xlPaperEnvelopeB6"
        Case xlPaperEnvelopeItaly: XlPaperSizeToString = "xlPaperEnvelopeItaly"
        Case xlPaperEnvelopeMonarch: XlPaperSizeToString = "xlPaperEnvelopeMonarch"
        Case xlPaperEnvelopePersonal: XlPaperSizeToString = "xlPaperEnvelopePersonal"
        Case xlPaperFanfoldUS: XlPaperSizeToString = "xlPaperFanfoldUS"
        Case xlPaperFanfoldStdGerman: XlPaperSizeToString = "xlPaperFanfoldStdGerman"
        Case xlPaperFanfoldLegalGerman: XlPaperSizeToString = "xlPaperFanfoldLegalGerman"
        Case xlPaperUser: XlPaperSizeToString = "xlPaperUser"
    End Select
End Function
