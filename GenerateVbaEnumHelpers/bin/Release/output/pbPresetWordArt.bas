Attribute VB_Name = "wpbPresetWordArt"
Function pbPresetWordArtFromString(value As String) As pbPresetWordArt
    If IsNumeric(value) Then
        pbPresetWordArtFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPresetWordArt1": pbPresetWordArtFromString = pbPresetWordArt1
        Case "pbPresetWordArt2": pbPresetWordArtFromString = pbPresetWordArt2
        Case "pbPresetWordArt3": pbPresetWordArtFromString = pbPresetWordArt3
        Case "pbPresetWordArt4": pbPresetWordArtFromString = pbPresetWordArt4
        Case "pbPresetWordArt5": pbPresetWordArtFromString = pbPresetWordArt5
        Case "pbPresetWordArt6": pbPresetWordArtFromString = pbPresetWordArt6
        Case "pbPresetWordArt7": pbPresetWordArtFromString = pbPresetWordArt7
        Case "pbPresetWordArt8": pbPresetWordArtFromString = pbPresetWordArt8
        Case "pbPresetWordArt9": pbPresetWordArtFromString = pbPresetWordArt9
        Case "pbPresetWordArt10": pbPresetWordArtFromString = pbPresetWordArt10
        Case "pbPresetWordArt11": pbPresetWordArtFromString = pbPresetWordArt11
        Case "pbPresetWordArt12": pbPresetWordArtFromString = pbPresetWordArt12
        Case "pbPresetWordArt13": pbPresetWordArtFromString = pbPresetWordArt13
        Case "pbPresetWordArt14": pbPresetWordArtFromString = pbPresetWordArt14
        Case "pbPresetWordArt15": pbPresetWordArtFromString = pbPresetWordArt15
        Case "pbPresetWordArt16": pbPresetWordArtFromString = pbPresetWordArt16
        Case "pbPresetWordArt17": pbPresetWordArtFromString = pbPresetWordArt17
        Case "pbPresetWordArt18": pbPresetWordArtFromString = pbPresetWordArt18
        Case "pbPresetWordArt19": pbPresetWordArtFromString = pbPresetWordArt19
        Case "pbPresetWordArt20": pbPresetWordArtFromString = pbPresetWordArt20
        Case "pbPresetWordArt21": pbPresetWordArtFromString = pbPresetWordArt21
        Case "pbPresetWordArt22": pbPresetWordArtFromString = pbPresetWordArt22
        Case "pbPresetWordArt23": pbPresetWordArtFromString = pbPresetWordArt23
        Case "pbPresetWordArt24": pbPresetWordArtFromString = pbPresetWordArt24
        Case "pbPresetWordArt25": pbPresetWordArtFromString = pbPresetWordArt25
        Case "pbPresetWordArt26": pbPresetWordArtFromString = pbPresetWordArt26
        Case "pbPresetWordArt27": pbPresetWordArtFromString = pbPresetWordArt27
        Case "pbPresetWordArt28": pbPresetWordArtFromString = pbPresetWordArt28
        Case "pbPresetWordArt29": pbPresetWordArtFromString = pbPresetWordArt29
        Case "pbPresetWordArt30": pbPresetWordArtFromString = pbPresetWordArt30
        Case "pbPresetWordArt31": pbPresetWordArtFromString = pbPresetWordArt31
        Case "pbPresetWordArt32": pbPresetWordArtFromString = pbPresetWordArt32
        Case "pbPresetWordArt33": pbPresetWordArtFromString = pbPresetWordArt33
        Case "pbPresetWordArt34": pbPresetWordArtFromString = pbPresetWordArt34
        Case "pbPresetWordArt35": pbPresetWordArtFromString = pbPresetWordArt35
        Case "pbPresetWordArt36": pbPresetWordArtFromString = pbPresetWordArt36
        Case "pbPresetWordArt37": pbPresetWordArtFromString = pbPresetWordArt37
        Case "pbPresetWordArt38": pbPresetWordArtFromString = pbPresetWordArt38
        Case "pbPresetWordArt39": pbPresetWordArtFromString = pbPresetWordArt39
        Case "pbPresetWordArt40": pbPresetWordArtFromString = pbPresetWordArt40
        Case "pbPresetWordArt41": pbPresetWordArtFromString = pbPresetWordArt41
        Case "pbPresetWordArt42": pbPresetWordArtFromString = pbPresetWordArt42
        Case "pbPresetWordArt43": pbPresetWordArtFromString = pbPresetWordArt43
        Case "pbPresetWordArt44": pbPresetWordArtFromString = pbPresetWordArt44
        Case "pbPresetWordArt45": pbPresetWordArtFromString = pbPresetWordArt45
        Case "pbPresetWordArt46": pbPresetWordArtFromString = pbPresetWordArt46
        Case "pbPresetWordArt47": pbPresetWordArtFromString = pbPresetWordArt47
        Case "pbPresetWordArt48": pbPresetWordArtFromString = pbPresetWordArt48
        Case "pbPresetWordArt49": pbPresetWordArtFromString = pbPresetWordArt49
        Case "pbPresetWordArt50": pbPresetWordArtFromString = pbPresetWordArt50
        Case "pbPresetWordArt51": pbPresetWordArtFromString = pbPresetWordArt51
        Case "pbPresetWordArt52": pbPresetWordArtFromString = pbPresetWordArt52
        Case "pbPresetWordArt53": pbPresetWordArtFromString = pbPresetWordArt53
        Case "pbPresetWordArt54": pbPresetWordArtFromString = pbPresetWordArt54
        Case "pbPresetWordArt55": pbPresetWordArtFromString = pbPresetWordArt55
        Case "pbPresetWordArt56": pbPresetWordArtFromString = pbPresetWordArt56
        Case "pbPresetWordArt57": pbPresetWordArtFromString = pbPresetWordArt57
        Case "pbPresetWordArt58": pbPresetWordArtFromString = pbPresetWordArt58
        Case "pbPresetWordArt59": pbPresetWordArtFromString = pbPresetWordArt59
        Case "pbPresetWordArt60": pbPresetWordArtFromString = pbPresetWordArt60
        Case "pbPresetWordArtMixed": pbPresetWordArtFromString = pbPresetWordArtMixed
    End Select
End Function

Function pbPresetWordArtToString(value As pbPresetWordArt) As String
    Select Case value
        Case pbPresetWordArt1: pbPresetWordArtToString = "pbPresetWordArt1"
        Case pbPresetWordArt2: pbPresetWordArtToString = "pbPresetWordArt2"
        Case pbPresetWordArt3: pbPresetWordArtToString = "pbPresetWordArt3"
        Case pbPresetWordArt4: pbPresetWordArtToString = "pbPresetWordArt4"
        Case pbPresetWordArt5: pbPresetWordArtToString = "pbPresetWordArt5"
        Case pbPresetWordArt6: pbPresetWordArtToString = "pbPresetWordArt6"
        Case pbPresetWordArt7: pbPresetWordArtToString = "pbPresetWordArt7"
        Case pbPresetWordArt8: pbPresetWordArtToString = "pbPresetWordArt8"
        Case pbPresetWordArt9: pbPresetWordArtToString = "pbPresetWordArt9"
        Case pbPresetWordArt10: pbPresetWordArtToString = "pbPresetWordArt10"
        Case pbPresetWordArt11: pbPresetWordArtToString = "pbPresetWordArt11"
        Case pbPresetWordArt12: pbPresetWordArtToString = "pbPresetWordArt12"
        Case pbPresetWordArt13: pbPresetWordArtToString = "pbPresetWordArt13"
        Case pbPresetWordArt14: pbPresetWordArtToString = "pbPresetWordArt14"
        Case pbPresetWordArt15: pbPresetWordArtToString = "pbPresetWordArt15"
        Case pbPresetWordArt16: pbPresetWordArtToString = "pbPresetWordArt16"
        Case pbPresetWordArt17: pbPresetWordArtToString = "pbPresetWordArt17"
        Case pbPresetWordArt18: pbPresetWordArtToString = "pbPresetWordArt18"
        Case pbPresetWordArt19: pbPresetWordArtToString = "pbPresetWordArt19"
        Case pbPresetWordArt20: pbPresetWordArtToString = "pbPresetWordArt20"
        Case pbPresetWordArt21: pbPresetWordArtToString = "pbPresetWordArt21"
        Case pbPresetWordArt22: pbPresetWordArtToString = "pbPresetWordArt22"
        Case pbPresetWordArt23: pbPresetWordArtToString = "pbPresetWordArt23"
        Case pbPresetWordArt24: pbPresetWordArtToString = "pbPresetWordArt24"
        Case pbPresetWordArt25: pbPresetWordArtToString = "pbPresetWordArt25"
        Case pbPresetWordArt26: pbPresetWordArtToString = "pbPresetWordArt26"
        Case pbPresetWordArt27: pbPresetWordArtToString = "pbPresetWordArt27"
        Case pbPresetWordArt28: pbPresetWordArtToString = "pbPresetWordArt28"
        Case pbPresetWordArt29: pbPresetWordArtToString = "pbPresetWordArt29"
        Case pbPresetWordArt30: pbPresetWordArtToString = "pbPresetWordArt30"
        Case pbPresetWordArt31: pbPresetWordArtToString = "pbPresetWordArt31"
        Case pbPresetWordArt32: pbPresetWordArtToString = "pbPresetWordArt32"
        Case pbPresetWordArt33: pbPresetWordArtToString = "pbPresetWordArt33"
        Case pbPresetWordArt34: pbPresetWordArtToString = "pbPresetWordArt34"
        Case pbPresetWordArt35: pbPresetWordArtToString = "pbPresetWordArt35"
        Case pbPresetWordArt36: pbPresetWordArtToString = "pbPresetWordArt36"
        Case pbPresetWordArt37: pbPresetWordArtToString = "pbPresetWordArt37"
        Case pbPresetWordArt38: pbPresetWordArtToString = "pbPresetWordArt38"
        Case pbPresetWordArt39: pbPresetWordArtToString = "pbPresetWordArt39"
        Case pbPresetWordArt40: pbPresetWordArtToString = "pbPresetWordArt40"
        Case pbPresetWordArt41: pbPresetWordArtToString = "pbPresetWordArt41"
        Case pbPresetWordArt42: pbPresetWordArtToString = "pbPresetWordArt42"
        Case pbPresetWordArt43: pbPresetWordArtToString = "pbPresetWordArt43"
        Case pbPresetWordArt44: pbPresetWordArtToString = "pbPresetWordArt44"
        Case pbPresetWordArt45: pbPresetWordArtToString = "pbPresetWordArt45"
        Case pbPresetWordArt46: pbPresetWordArtToString = "pbPresetWordArt46"
        Case pbPresetWordArt47: pbPresetWordArtToString = "pbPresetWordArt47"
        Case pbPresetWordArt48: pbPresetWordArtToString = "pbPresetWordArt48"
        Case pbPresetWordArt49: pbPresetWordArtToString = "pbPresetWordArt49"
        Case pbPresetWordArt50: pbPresetWordArtToString = "pbPresetWordArt50"
        Case pbPresetWordArt51: pbPresetWordArtToString = "pbPresetWordArt51"
        Case pbPresetWordArt52: pbPresetWordArtToString = "pbPresetWordArt52"
        Case pbPresetWordArt53: pbPresetWordArtToString = "pbPresetWordArt53"
        Case pbPresetWordArt54: pbPresetWordArtToString = "pbPresetWordArt54"
        Case pbPresetWordArt55: pbPresetWordArtToString = "pbPresetWordArt55"
        Case pbPresetWordArt56: pbPresetWordArtToString = "pbPresetWordArt56"
        Case pbPresetWordArt57: pbPresetWordArtToString = "pbPresetWordArt57"
        Case pbPresetWordArt58: pbPresetWordArtToString = "pbPresetWordArt58"
        Case pbPresetWordArt59: pbPresetWordArtToString = "pbPresetWordArt59"
        Case pbPresetWordArt60: pbPresetWordArtToString = "pbPresetWordArt60"
        Case pbPresetWordArtMixed: pbPresetWordArtToString = "pbPresetWordArtMixed"
    End Select
End Function
