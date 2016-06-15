Attribute VB_Name = "wMsoWarpFormat"
Function MsoWarpFormatFromString(value As String) As MsoWarpFormat
    If IsNumeric(value) Then
        MsoWarpFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoWarpFormat1": MsoWarpFormatFromString = msoWarpFormat1
        Case "msoWarpFormat2": MsoWarpFormatFromString = msoWarpFormat2
        Case "msoWarpFormat3": MsoWarpFormatFromString = msoWarpFormat3
        Case "msoWarpFormat4": MsoWarpFormatFromString = msoWarpFormat4
        Case "msoWarpFormat5": MsoWarpFormatFromString = msoWarpFormat5
        Case "msoWarpFormat6": MsoWarpFormatFromString = msoWarpFormat6
        Case "msoWarpFormat7": MsoWarpFormatFromString = msoWarpFormat7
        Case "msoWarpFormat8": MsoWarpFormatFromString = msoWarpFormat8
        Case "msoWarpFormat9": MsoWarpFormatFromString = msoWarpFormat9
        Case "msoWarpFormat10": MsoWarpFormatFromString = msoWarpFormat10
        Case "msoWarpFormat11": MsoWarpFormatFromString = msoWarpFormat11
        Case "msoWarpFormat12": MsoWarpFormatFromString = msoWarpFormat12
        Case "msoWarpFormat13": MsoWarpFormatFromString = msoWarpFormat13
        Case "msoWarpFormat14": MsoWarpFormatFromString = msoWarpFormat14
        Case "msoWarpFormat15": MsoWarpFormatFromString = msoWarpFormat15
        Case "msoWarpFormat16": MsoWarpFormatFromString = msoWarpFormat16
        Case "msoWarpFormat17": MsoWarpFormatFromString = msoWarpFormat17
        Case "msoWarpFormat18": MsoWarpFormatFromString = msoWarpFormat18
        Case "msoWarpFormat19": MsoWarpFormatFromString = msoWarpFormat19
        Case "msoWarpFormat20": MsoWarpFormatFromString = msoWarpFormat20
        Case "msoWarpFormat21": MsoWarpFormatFromString = msoWarpFormat21
        Case "msoWarpFormat22": MsoWarpFormatFromString = msoWarpFormat22
        Case "msoWarpFormat23": MsoWarpFormatFromString = msoWarpFormat23
        Case "msoWarpFormat24": MsoWarpFormatFromString = msoWarpFormat24
        Case "msoWarpFormat25": MsoWarpFormatFromString = msoWarpFormat25
        Case "msoWarpFormat26": MsoWarpFormatFromString = msoWarpFormat26
        Case "msoWarpFormat27": MsoWarpFormatFromString = msoWarpFormat27
        Case "msoWarpFormat28": MsoWarpFormatFromString = msoWarpFormat28
        Case "msoWarpFormat29": MsoWarpFormatFromString = msoWarpFormat29
        Case "msoWarpFormat30": MsoWarpFormatFromString = msoWarpFormat30
        Case "msoWarpFormat31": MsoWarpFormatFromString = msoWarpFormat31
        Case "msoWarpFormat32": MsoWarpFormatFromString = msoWarpFormat32
        Case "msoWarpFormat33": MsoWarpFormatFromString = msoWarpFormat33
        Case "msoWarpFormat34": MsoWarpFormatFromString = msoWarpFormat34
        Case "msoWarpFormat35": MsoWarpFormatFromString = msoWarpFormat35
        Case "msoWarpFormat36": MsoWarpFormatFromString = msoWarpFormat36
        Case "msoWarpFormatMixed": MsoWarpFormatFromString = msoWarpFormatMixed
    End Select
End Function

Function MsoWarpFormatToString(value As MsoWarpFormat) As String
    Select Case value
        Case msoWarpFormat1: MsoWarpFormatToString = "msoWarpFormat1"
        Case msoWarpFormat2: MsoWarpFormatToString = "msoWarpFormat2"
        Case msoWarpFormat3: MsoWarpFormatToString = "msoWarpFormat3"
        Case msoWarpFormat4: MsoWarpFormatToString = "msoWarpFormat4"
        Case msoWarpFormat5: MsoWarpFormatToString = "msoWarpFormat5"
        Case msoWarpFormat6: MsoWarpFormatToString = "msoWarpFormat6"
        Case msoWarpFormat7: MsoWarpFormatToString = "msoWarpFormat7"
        Case msoWarpFormat8: MsoWarpFormatToString = "msoWarpFormat8"
        Case msoWarpFormat9: MsoWarpFormatToString = "msoWarpFormat9"
        Case msoWarpFormat10: MsoWarpFormatToString = "msoWarpFormat10"
        Case msoWarpFormat11: MsoWarpFormatToString = "msoWarpFormat11"
        Case msoWarpFormat12: MsoWarpFormatToString = "msoWarpFormat12"
        Case msoWarpFormat13: MsoWarpFormatToString = "msoWarpFormat13"
        Case msoWarpFormat14: MsoWarpFormatToString = "msoWarpFormat14"
        Case msoWarpFormat15: MsoWarpFormatToString = "msoWarpFormat15"
        Case msoWarpFormat16: MsoWarpFormatToString = "msoWarpFormat16"
        Case msoWarpFormat17: MsoWarpFormatToString = "msoWarpFormat17"
        Case msoWarpFormat18: MsoWarpFormatToString = "msoWarpFormat18"
        Case msoWarpFormat19: MsoWarpFormatToString = "msoWarpFormat19"
        Case msoWarpFormat20: MsoWarpFormatToString = "msoWarpFormat20"
        Case msoWarpFormat21: MsoWarpFormatToString = "msoWarpFormat21"
        Case msoWarpFormat22: MsoWarpFormatToString = "msoWarpFormat22"
        Case msoWarpFormat23: MsoWarpFormatToString = "msoWarpFormat23"
        Case msoWarpFormat24: MsoWarpFormatToString = "msoWarpFormat24"
        Case msoWarpFormat25: MsoWarpFormatToString = "msoWarpFormat25"
        Case msoWarpFormat26: MsoWarpFormatToString = "msoWarpFormat26"
        Case msoWarpFormat27: MsoWarpFormatToString = "msoWarpFormat27"
        Case msoWarpFormat28: MsoWarpFormatToString = "msoWarpFormat28"
        Case msoWarpFormat29: MsoWarpFormatToString = "msoWarpFormat29"
        Case msoWarpFormat30: MsoWarpFormatToString = "msoWarpFormat30"
        Case msoWarpFormat31: MsoWarpFormatToString = "msoWarpFormat31"
        Case msoWarpFormat32: MsoWarpFormatToString = "msoWarpFormat32"
        Case msoWarpFormat33: MsoWarpFormatToString = "msoWarpFormat33"
        Case msoWarpFormat34: MsoWarpFormatToString = "msoWarpFormat34"
        Case msoWarpFormat35: MsoWarpFormatToString = "msoWarpFormat35"
        Case msoWarpFormat36: MsoWarpFormatToString = "msoWarpFormat36"
        Case msoWarpFormatMixed: MsoWarpFormatToString = "msoWarpFormatMixed"
    End Select
End Function
