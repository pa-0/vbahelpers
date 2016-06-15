Attribute VB_Name = "wMsoPresetTextEffect"
Function MsoPresetTextEffectFromString(value As String) As MsoPresetTextEffect
    If IsNumeric(value) Then
        MsoPresetTextEffectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTextEffect1": MsoPresetTextEffectFromString = msoTextEffect1
        Case "msoTextEffect2": MsoPresetTextEffectFromString = msoTextEffect2
        Case "msoTextEffect3": MsoPresetTextEffectFromString = msoTextEffect3
        Case "msoTextEffect4": MsoPresetTextEffectFromString = msoTextEffect4
        Case "msoTextEffect5": MsoPresetTextEffectFromString = msoTextEffect5
        Case "msoTextEffect6": MsoPresetTextEffectFromString = msoTextEffect6
        Case "msoTextEffect7": MsoPresetTextEffectFromString = msoTextEffect7
        Case "msoTextEffect8": MsoPresetTextEffectFromString = msoTextEffect8
        Case "msoTextEffect9": MsoPresetTextEffectFromString = msoTextEffect9
        Case "msoTextEffect10": MsoPresetTextEffectFromString = msoTextEffect10
        Case "msoTextEffect11": MsoPresetTextEffectFromString = msoTextEffect11
        Case "msoTextEffect12": MsoPresetTextEffectFromString = msoTextEffect12
        Case "msoTextEffect13": MsoPresetTextEffectFromString = msoTextEffect13
        Case "msoTextEffect14": MsoPresetTextEffectFromString = msoTextEffect14
        Case "msoTextEffect15": MsoPresetTextEffectFromString = msoTextEffect15
        Case "msoTextEffect16": MsoPresetTextEffectFromString = msoTextEffect16
        Case "msoTextEffect17": MsoPresetTextEffectFromString = msoTextEffect17
        Case "msoTextEffect18": MsoPresetTextEffectFromString = msoTextEffect18
        Case "msoTextEffect19": MsoPresetTextEffectFromString = msoTextEffect19
        Case "msoTextEffect20": MsoPresetTextEffectFromString = msoTextEffect20
        Case "msoTextEffect21": MsoPresetTextEffectFromString = msoTextEffect21
        Case "msoTextEffect22": MsoPresetTextEffectFromString = msoTextEffect22
        Case "msoTextEffect23": MsoPresetTextEffectFromString = msoTextEffect23
        Case "msoTextEffect24": MsoPresetTextEffectFromString = msoTextEffect24
        Case "msoTextEffect25": MsoPresetTextEffectFromString = msoTextEffect25
        Case "msoTextEffect26": MsoPresetTextEffectFromString = msoTextEffect26
        Case "msoTextEffect27": MsoPresetTextEffectFromString = msoTextEffect27
        Case "msoTextEffect28": MsoPresetTextEffectFromString = msoTextEffect28
        Case "msoTextEffect29": MsoPresetTextEffectFromString = msoTextEffect29
        Case "msoTextEffect30": MsoPresetTextEffectFromString = msoTextEffect30
        Case "msoTextEffectMixed": MsoPresetTextEffectFromString = msoTextEffectMixed
    End Select
End Function

Function MsoPresetTextEffectToString(value As MsoPresetTextEffect) As String
    Select Case value
        Case msoTextEffect1: MsoPresetTextEffectToString = "msoTextEffect1"
        Case msoTextEffect2: MsoPresetTextEffectToString = "msoTextEffect2"
        Case msoTextEffect3: MsoPresetTextEffectToString = "msoTextEffect3"
        Case msoTextEffect4: MsoPresetTextEffectToString = "msoTextEffect4"
        Case msoTextEffect5: MsoPresetTextEffectToString = "msoTextEffect5"
        Case msoTextEffect6: MsoPresetTextEffectToString = "msoTextEffect6"
        Case msoTextEffect7: MsoPresetTextEffectToString = "msoTextEffect7"
        Case msoTextEffect8: MsoPresetTextEffectToString = "msoTextEffect8"
        Case msoTextEffect9: MsoPresetTextEffectToString = "msoTextEffect9"
        Case msoTextEffect10: MsoPresetTextEffectToString = "msoTextEffect10"
        Case msoTextEffect11: MsoPresetTextEffectToString = "msoTextEffect11"
        Case msoTextEffect12: MsoPresetTextEffectToString = "msoTextEffect12"
        Case msoTextEffect13: MsoPresetTextEffectToString = "msoTextEffect13"
        Case msoTextEffect14: MsoPresetTextEffectToString = "msoTextEffect14"
        Case msoTextEffect15: MsoPresetTextEffectToString = "msoTextEffect15"
        Case msoTextEffect16: MsoPresetTextEffectToString = "msoTextEffect16"
        Case msoTextEffect17: MsoPresetTextEffectToString = "msoTextEffect17"
        Case msoTextEffect18: MsoPresetTextEffectToString = "msoTextEffect18"
        Case msoTextEffect19: MsoPresetTextEffectToString = "msoTextEffect19"
        Case msoTextEffect20: MsoPresetTextEffectToString = "msoTextEffect20"
        Case msoTextEffect21: MsoPresetTextEffectToString = "msoTextEffect21"
        Case msoTextEffect22: MsoPresetTextEffectToString = "msoTextEffect22"
        Case msoTextEffect23: MsoPresetTextEffectToString = "msoTextEffect23"
        Case msoTextEffect24: MsoPresetTextEffectToString = "msoTextEffect24"
        Case msoTextEffect25: MsoPresetTextEffectToString = "msoTextEffect25"
        Case msoTextEffect26: MsoPresetTextEffectToString = "msoTextEffect26"
        Case msoTextEffect27: MsoPresetTextEffectToString = "msoTextEffect27"
        Case msoTextEffect28: MsoPresetTextEffectToString = "msoTextEffect28"
        Case msoTextEffect29: MsoPresetTextEffectToString = "msoTextEffect29"
        Case msoTextEffect30: MsoPresetTextEffectToString = "msoTextEffect30"
        Case msoTextEffectMixed: MsoPresetTextEffectToString = "msoTextEffectMixed"
    End Select
End Function
