Attribute VB_Name = "wWdCountry"
Function WdCountryFromString(value As String) As WdCountry
    If IsNumeric(value) Then
        WdCountryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdUS": WdCountryFromString = wdUS
        Case "wdCanada": WdCountryFromString = wdCanada
        Case "wdLatinAmerica": WdCountryFromString = wdLatinAmerica
        Case "wdNetherlands": WdCountryFromString = wdNetherlands
        Case "wdFrance": WdCountryFromString = wdFrance
        Case "wdSpain": WdCountryFromString = wdSpain
        Case "wdItaly": WdCountryFromString = wdItaly
        Case "wdUK": WdCountryFromString = wdUK
        Case "wdDenmark": WdCountryFromString = wdDenmark
        Case "wdSweden": WdCountryFromString = wdSweden
        Case "wdNorway": WdCountryFromString = wdNorway
        Case "wdGermany": WdCountryFromString = wdGermany
        Case "wdPeru": WdCountryFromString = wdPeru
        Case "wdMexico": WdCountryFromString = wdMexico
        Case "wdArgentina": WdCountryFromString = wdArgentina
        Case "wdBrazil": WdCountryFromString = wdBrazil
        Case "wdChile": WdCountryFromString = wdChile
        Case "wdVenezuela": WdCountryFromString = wdVenezuela
        Case "wdJapan": WdCountryFromString = wdJapan
        Case "wdKorea": WdCountryFromString = wdKorea
        Case "wdChina": WdCountryFromString = wdChina
        Case "wdIceland": WdCountryFromString = wdIceland
        Case "wdFinland": WdCountryFromString = wdFinland
        Case "wdTaiwan": WdCountryFromString = wdTaiwan
    End Select
End Function

Function WdCountryToString(value As WdCountry) As String
    Select Case value
        Case wdUS: WdCountryToString = "wdUS"
        Case wdCanada: WdCountryToString = "wdCanada"
        Case wdLatinAmerica: WdCountryToString = "wdLatinAmerica"
        Case wdNetherlands: WdCountryToString = "wdNetherlands"
        Case wdFrance: WdCountryToString = "wdFrance"
        Case wdSpain: WdCountryToString = "wdSpain"
        Case wdItaly: WdCountryToString = "wdItaly"
        Case wdUK: WdCountryToString = "wdUK"
        Case wdDenmark: WdCountryToString = "wdDenmark"
        Case wdSweden: WdCountryToString = "wdSweden"
        Case wdNorway: WdCountryToString = "wdNorway"
        Case wdGermany: WdCountryToString = "wdGermany"
        Case wdPeru: WdCountryToString = "wdPeru"
        Case wdMexico: WdCountryToString = "wdMexico"
        Case wdArgentina: WdCountryToString = "wdArgentina"
        Case wdBrazil: WdCountryToString = "wdBrazil"
        Case wdChile: WdCountryToString = "wdChile"
        Case wdVenezuela: WdCountryToString = "wdVenezuela"
        Case wdJapan: WdCountryToString = "wdJapan"
        Case wdKorea: WdCountryToString = "wdKorea"
        Case wdChina: WdCountryToString = "wdChina"
        Case wdIceland: WdCountryToString = "wdIceland"
        Case wdFinland: WdCountryToString = "wdFinland"
        Case wdTaiwan: WdCountryToString = "wdTaiwan"
    End Select
End Function
