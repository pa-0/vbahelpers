Attribute VB_Name = "wMsoShadowType"
Function MsoShadowTypeFromString(value As String) As MsoShadowType
    If IsNumeric(value) Then
        MsoShadowTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoShadow1": MsoShadowTypeFromString = msoShadow1
        Case "msoShadow2": MsoShadowTypeFromString = msoShadow2
        Case "msoShadow3": MsoShadowTypeFromString = msoShadow3
        Case "msoShadow4": MsoShadowTypeFromString = msoShadow4
        Case "msoShadow5": MsoShadowTypeFromString = msoShadow5
        Case "msoShadow6": MsoShadowTypeFromString = msoShadow6
        Case "msoShadow7": MsoShadowTypeFromString = msoShadow7
        Case "msoShadow8": MsoShadowTypeFromString = msoShadow8
        Case "msoShadow9": MsoShadowTypeFromString = msoShadow9
        Case "msoShadow10": MsoShadowTypeFromString = msoShadow10
        Case "msoShadow11": MsoShadowTypeFromString = msoShadow11
        Case "msoShadow12": MsoShadowTypeFromString = msoShadow12
        Case "msoShadow13": MsoShadowTypeFromString = msoShadow13
        Case "msoShadow14": MsoShadowTypeFromString = msoShadow14
        Case "msoShadow15": MsoShadowTypeFromString = msoShadow15
        Case "msoShadow16": MsoShadowTypeFromString = msoShadow16
        Case "msoShadow17": MsoShadowTypeFromString = msoShadow17
        Case "msoShadow18": MsoShadowTypeFromString = msoShadow18
        Case "msoShadow19": MsoShadowTypeFromString = msoShadow19
        Case "msoShadow20": MsoShadowTypeFromString = msoShadow20
        Case "msoShadow21": MsoShadowTypeFromString = msoShadow21
        Case "msoShadow22": MsoShadowTypeFromString = msoShadow22
        Case "msoShadow23": MsoShadowTypeFromString = msoShadow23
        Case "msoShadow24": MsoShadowTypeFromString = msoShadow24
        Case "msoShadow25": MsoShadowTypeFromString = msoShadow25
        Case "msoShadow26": MsoShadowTypeFromString = msoShadow26
        Case "msoShadow27": MsoShadowTypeFromString = msoShadow27
        Case "msoShadow28": MsoShadowTypeFromString = msoShadow28
        Case "msoShadow29": MsoShadowTypeFromString = msoShadow29
        Case "msoShadow30": MsoShadowTypeFromString = msoShadow30
        Case "msoShadow31": MsoShadowTypeFromString = msoShadow31
        Case "msoShadow32": MsoShadowTypeFromString = msoShadow32
        Case "msoShadow33": MsoShadowTypeFromString = msoShadow33
        Case "msoShadow34": MsoShadowTypeFromString = msoShadow34
        Case "msoShadow35": MsoShadowTypeFromString = msoShadow35
        Case "msoShadow36": MsoShadowTypeFromString = msoShadow36
        Case "msoShadow37": MsoShadowTypeFromString = msoShadow37
        Case "msoShadow38": MsoShadowTypeFromString = msoShadow38
        Case "msoShadow39": MsoShadowTypeFromString = msoShadow39
        Case "msoShadow40": MsoShadowTypeFromString = msoShadow40
        Case "msoShadow41": MsoShadowTypeFromString = msoShadow41
        Case "msoShadow42": MsoShadowTypeFromString = msoShadow42
        Case "msoShadow43": MsoShadowTypeFromString = msoShadow43
        Case "msoShadowMixed": MsoShadowTypeFromString = msoShadowMixed
    End Select
End Function

Function MsoShadowTypeToString(value As MsoShadowType) As String
    Select Case value
        Case msoShadow1: MsoShadowTypeToString = "msoShadow1"
        Case msoShadow2: MsoShadowTypeToString = "msoShadow2"
        Case msoShadow3: MsoShadowTypeToString = "msoShadow3"
        Case msoShadow4: MsoShadowTypeToString = "msoShadow4"
        Case msoShadow5: MsoShadowTypeToString = "msoShadow5"
        Case msoShadow6: MsoShadowTypeToString = "msoShadow6"
        Case msoShadow7: MsoShadowTypeToString = "msoShadow7"
        Case msoShadow8: MsoShadowTypeToString = "msoShadow8"
        Case msoShadow9: MsoShadowTypeToString = "msoShadow9"
        Case msoShadow10: MsoShadowTypeToString = "msoShadow10"
        Case msoShadow11: MsoShadowTypeToString = "msoShadow11"
        Case msoShadow12: MsoShadowTypeToString = "msoShadow12"
        Case msoShadow13: MsoShadowTypeToString = "msoShadow13"
        Case msoShadow14: MsoShadowTypeToString = "msoShadow14"
        Case msoShadow15: MsoShadowTypeToString = "msoShadow15"
        Case msoShadow16: MsoShadowTypeToString = "msoShadow16"
        Case msoShadow17: MsoShadowTypeToString = "msoShadow17"
        Case msoShadow18: MsoShadowTypeToString = "msoShadow18"
        Case msoShadow19: MsoShadowTypeToString = "msoShadow19"
        Case msoShadow20: MsoShadowTypeToString = "msoShadow20"
        Case msoShadow21: MsoShadowTypeToString = "msoShadow21"
        Case msoShadow22: MsoShadowTypeToString = "msoShadow22"
        Case msoShadow23: MsoShadowTypeToString = "msoShadow23"
        Case msoShadow24: MsoShadowTypeToString = "msoShadow24"
        Case msoShadow25: MsoShadowTypeToString = "msoShadow25"
        Case msoShadow26: MsoShadowTypeToString = "msoShadow26"
        Case msoShadow27: MsoShadowTypeToString = "msoShadow27"
        Case msoShadow28: MsoShadowTypeToString = "msoShadow28"
        Case msoShadow29: MsoShadowTypeToString = "msoShadow29"
        Case msoShadow30: MsoShadowTypeToString = "msoShadow30"
        Case msoShadow31: MsoShadowTypeToString = "msoShadow31"
        Case msoShadow32: MsoShadowTypeToString = "msoShadow32"
        Case msoShadow33: MsoShadowTypeToString = "msoShadow33"
        Case msoShadow34: MsoShadowTypeToString = "msoShadow34"
        Case msoShadow35: MsoShadowTypeToString = "msoShadow35"
        Case msoShadow36: MsoShadowTypeToString = "msoShadow36"
        Case msoShadow37: MsoShadowTypeToString = "msoShadow37"
        Case msoShadow38: MsoShadowTypeToString = "msoShadow38"
        Case msoShadow39: MsoShadowTypeToString = "msoShadow39"
        Case msoShadow40: MsoShadowTypeToString = "msoShadow40"
        Case msoShadow41: MsoShadowTypeToString = "msoShadow41"
        Case msoShadow42: MsoShadowTypeToString = "msoShadow42"
        Case msoShadow43: MsoShadowTypeToString = "msoShadow43"
        Case msoShadowMixed: MsoShadowTypeToString = "msoShadowMixed"
    End Select
End Function
