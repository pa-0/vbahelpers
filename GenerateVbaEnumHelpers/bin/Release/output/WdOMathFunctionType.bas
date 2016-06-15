Attribute VB_Name = "wWdOMathFunctionType"
Function WdOMathFunctionTypeFromString(value As String) As WdOMathFunctionType
    If IsNumeric(value) Then
        WdOMathFunctionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathFunctionAcc": WdOMathFunctionTypeFromString = wdOMathFunctionAcc
        Case "wdOMathFunctionBar": WdOMathFunctionTypeFromString = wdOMathFunctionBar
        Case "wdOMathFunctionBox": WdOMathFunctionTypeFromString = wdOMathFunctionBox
        Case "wdOMathFunctionBorderBox": WdOMathFunctionTypeFromString = wdOMathFunctionBorderBox
        Case "wdOMathFunctionDelim": WdOMathFunctionTypeFromString = wdOMathFunctionDelim
        Case "wdOMathFunctionEqArray": WdOMathFunctionTypeFromString = wdOMathFunctionEqArray
        Case "wdOMathFunctionFrac": WdOMathFunctionTypeFromString = wdOMathFunctionFrac
        Case "wdOMathFunctionFunc": WdOMathFunctionTypeFromString = wdOMathFunctionFunc
        Case "wdOMathFunctionGroupChar": WdOMathFunctionTypeFromString = wdOMathFunctionGroupChar
        Case "wdOMathFunctionLimLow": WdOMathFunctionTypeFromString = wdOMathFunctionLimLow
        Case "wdOMathFunctionLimUpp": WdOMathFunctionTypeFromString = wdOMathFunctionLimUpp
        Case "wdOMathFunctionMat": WdOMathFunctionTypeFromString = wdOMathFunctionMat
        Case "wdOMathFunctionNary": WdOMathFunctionTypeFromString = wdOMathFunctionNary
        Case "wdOMathFunctionPhantom": WdOMathFunctionTypeFromString = wdOMathFunctionPhantom
        Case "wdOMathFunctionScrPre": WdOMathFunctionTypeFromString = wdOMathFunctionScrPre
        Case "wdOMathFunctionRad": WdOMathFunctionTypeFromString = wdOMathFunctionRad
        Case "wdOMathFunctionScrSub": WdOMathFunctionTypeFromString = wdOMathFunctionScrSub
        Case "wdOMathFunctionScrSubSup": WdOMathFunctionTypeFromString = wdOMathFunctionScrSubSup
        Case "wdOMathFunctionScrSup": WdOMathFunctionTypeFromString = wdOMathFunctionScrSup
        Case "wdOMathFunctionText": WdOMathFunctionTypeFromString = wdOMathFunctionText
        Case "wdOMathFunctionNormalText": WdOMathFunctionTypeFromString = wdOMathFunctionNormalText
        Case "wdOMathFunctionLiteralText": WdOMathFunctionTypeFromString = wdOMathFunctionLiteralText
    End Select
End Function

Function WdOMathFunctionTypeToString(value As WdOMathFunctionType) As String
    Select Case value
        Case wdOMathFunctionAcc: WdOMathFunctionTypeToString = "wdOMathFunctionAcc"
        Case wdOMathFunctionBar: WdOMathFunctionTypeToString = "wdOMathFunctionBar"
        Case wdOMathFunctionBox: WdOMathFunctionTypeToString = "wdOMathFunctionBox"
        Case wdOMathFunctionBorderBox: WdOMathFunctionTypeToString = "wdOMathFunctionBorderBox"
        Case wdOMathFunctionDelim: WdOMathFunctionTypeToString = "wdOMathFunctionDelim"
        Case wdOMathFunctionEqArray: WdOMathFunctionTypeToString = "wdOMathFunctionEqArray"
        Case wdOMathFunctionFrac: WdOMathFunctionTypeToString = "wdOMathFunctionFrac"
        Case wdOMathFunctionFunc: WdOMathFunctionTypeToString = "wdOMathFunctionFunc"
        Case wdOMathFunctionGroupChar: WdOMathFunctionTypeToString = "wdOMathFunctionGroupChar"
        Case wdOMathFunctionLimLow: WdOMathFunctionTypeToString = "wdOMathFunctionLimLow"
        Case wdOMathFunctionLimUpp: WdOMathFunctionTypeToString = "wdOMathFunctionLimUpp"
        Case wdOMathFunctionMat: WdOMathFunctionTypeToString = "wdOMathFunctionMat"
        Case wdOMathFunctionNary: WdOMathFunctionTypeToString = "wdOMathFunctionNary"
        Case wdOMathFunctionPhantom: WdOMathFunctionTypeToString = "wdOMathFunctionPhantom"
        Case wdOMathFunctionScrPre: WdOMathFunctionTypeToString = "wdOMathFunctionScrPre"
        Case wdOMathFunctionRad: WdOMathFunctionTypeToString = "wdOMathFunctionRad"
        Case wdOMathFunctionScrSub: WdOMathFunctionTypeToString = "wdOMathFunctionScrSub"
        Case wdOMathFunctionScrSubSup: WdOMathFunctionTypeToString = "wdOMathFunctionScrSubSup"
        Case wdOMathFunctionScrSup: WdOMathFunctionTypeToString = "wdOMathFunctionScrSup"
        Case wdOMathFunctionText: WdOMathFunctionTypeToString = "wdOMathFunctionText"
        Case wdOMathFunctionNormalText: WdOMathFunctionTypeToString = "wdOMathFunctionNormalText"
        Case wdOMathFunctionLiteralText: WdOMathFunctionTypeToString = "wdOMathFunctionLiteralText"
    End Select
End Function
