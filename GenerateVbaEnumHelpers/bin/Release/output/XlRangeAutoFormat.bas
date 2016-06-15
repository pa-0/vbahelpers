Attribute VB_Name = "wXlRangeAutoFormat"
Function XlRangeAutoFormatFromString(value As String) As XlRangeAutoFormat
    If IsNumeric(value) Then
        XlRangeAutoFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlRangeAutoFormatClassic1": XlRangeAutoFormatFromString = xlRangeAutoFormatClassic1
        Case "xlRangeAutoFormatClassic2": XlRangeAutoFormatFromString = xlRangeAutoFormatClassic2
        Case "xlRangeAutoFormatClassic3": XlRangeAutoFormatFromString = xlRangeAutoFormatClassic3
        Case "xlRangeAutoFormatAccounting1": XlRangeAutoFormatFromString = xlRangeAutoFormatAccounting1
        Case "xlRangeAutoFormatAccounting2": XlRangeAutoFormatFromString = xlRangeAutoFormatAccounting2
        Case "xlRangeAutoFormatAccounting3": XlRangeAutoFormatFromString = xlRangeAutoFormatAccounting3
        Case "xlRangeAutoFormatColor1": XlRangeAutoFormatFromString = xlRangeAutoFormatColor1
        Case "xlRangeAutoFormatColor2": XlRangeAutoFormatFromString = xlRangeAutoFormatColor2
        Case "xlRangeAutoFormatColor3": XlRangeAutoFormatFromString = xlRangeAutoFormatColor3
        Case "xlRangeAutoFormatList1": XlRangeAutoFormatFromString = xlRangeAutoFormatList1
        Case "xlRangeAutoFormatList2": XlRangeAutoFormatFromString = xlRangeAutoFormatList2
        Case "xlRangeAutoFormatList3": XlRangeAutoFormatFromString = xlRangeAutoFormatList3
        Case "xlRangeAutoFormat3DEffects1": XlRangeAutoFormatFromString = xlRangeAutoFormat3DEffects1
        Case "xlRangeAutoFormat3DEffects2": XlRangeAutoFormatFromString = xlRangeAutoFormat3DEffects2
        Case "xlRangeAutoFormatLocalFormat1": XlRangeAutoFormatFromString = xlRangeAutoFormatLocalFormat1
        Case "xlRangeAutoFormatLocalFormat2": XlRangeAutoFormatFromString = xlRangeAutoFormatLocalFormat2
        Case "xlRangeAutoFormatAccounting4": XlRangeAutoFormatFromString = xlRangeAutoFormatAccounting4
        Case "xlRangeAutoFormatLocalFormat3": XlRangeAutoFormatFromString = xlRangeAutoFormatLocalFormat3
        Case "xlRangeAutoFormatLocalFormat4": XlRangeAutoFormatFromString = xlRangeAutoFormatLocalFormat4
        Case "xlRangeAutoFormatReport1": XlRangeAutoFormatFromString = xlRangeAutoFormatReport1
        Case "xlRangeAutoFormatReport2": XlRangeAutoFormatFromString = xlRangeAutoFormatReport2
        Case "xlRangeAutoFormatReport3": XlRangeAutoFormatFromString = xlRangeAutoFormatReport3
        Case "xlRangeAutoFormatReport4": XlRangeAutoFormatFromString = xlRangeAutoFormatReport4
        Case "xlRangeAutoFormatReport5": XlRangeAutoFormatFromString = xlRangeAutoFormatReport5
        Case "xlRangeAutoFormatReport6": XlRangeAutoFormatFromString = xlRangeAutoFormatReport6
        Case "xlRangeAutoFormatReport7": XlRangeAutoFormatFromString = xlRangeAutoFormatReport7
        Case "xlRangeAutoFormatReport8": XlRangeAutoFormatFromString = xlRangeAutoFormatReport8
        Case "xlRangeAutoFormatReport9": XlRangeAutoFormatFromString = xlRangeAutoFormatReport9
        Case "xlRangeAutoFormatReport10": XlRangeAutoFormatFromString = xlRangeAutoFormatReport10
        Case "xlRangeAutoFormatClassicPivotTable": XlRangeAutoFormatFromString = xlRangeAutoFormatClassicPivotTable
        Case "xlRangeAutoFormatTable1": XlRangeAutoFormatFromString = xlRangeAutoFormatTable1
        Case "xlRangeAutoFormatTable2": XlRangeAutoFormatFromString = xlRangeAutoFormatTable2
        Case "xlRangeAutoFormatTable3": XlRangeAutoFormatFromString = xlRangeAutoFormatTable3
        Case "xlRangeAutoFormatTable4": XlRangeAutoFormatFromString = xlRangeAutoFormatTable4
        Case "xlRangeAutoFormatTable5": XlRangeAutoFormatFromString = xlRangeAutoFormatTable5
        Case "xlRangeAutoFormatTable6": XlRangeAutoFormatFromString = xlRangeAutoFormatTable6
        Case "xlRangeAutoFormatTable7": XlRangeAutoFormatFromString = xlRangeAutoFormatTable7
        Case "xlRangeAutoFormatTable8": XlRangeAutoFormatFromString = xlRangeAutoFormatTable8
        Case "xlRangeAutoFormatTable9": XlRangeAutoFormatFromString = xlRangeAutoFormatTable9
        Case "xlRangeAutoFormatTable10": XlRangeAutoFormatFromString = xlRangeAutoFormatTable10
        Case "xlRangeAutoFormatPTNone": XlRangeAutoFormatFromString = xlRangeAutoFormatPTNone
        Case "xlRangeAutoFormatSimple": XlRangeAutoFormatFromString = xlRangeAutoFormatSimple
        Case "xlRangeAutoFormatNone": XlRangeAutoFormatFromString = xlRangeAutoFormatNone
    End Select
End Function

Function XlRangeAutoFormatToString(value As XlRangeAutoFormat) As String
    Select Case value
        Case xlRangeAutoFormatClassic1: XlRangeAutoFormatToString = "xlRangeAutoFormatClassic1"
        Case xlRangeAutoFormatClassic2: XlRangeAutoFormatToString = "xlRangeAutoFormatClassic2"
        Case xlRangeAutoFormatClassic3: XlRangeAutoFormatToString = "xlRangeAutoFormatClassic3"
        Case xlRangeAutoFormatAccounting1: XlRangeAutoFormatToString = "xlRangeAutoFormatAccounting1"
        Case xlRangeAutoFormatAccounting2: XlRangeAutoFormatToString = "xlRangeAutoFormatAccounting2"
        Case xlRangeAutoFormatAccounting3: XlRangeAutoFormatToString = "xlRangeAutoFormatAccounting3"
        Case xlRangeAutoFormatColor1: XlRangeAutoFormatToString = "xlRangeAutoFormatColor1"
        Case xlRangeAutoFormatColor2: XlRangeAutoFormatToString = "xlRangeAutoFormatColor2"
        Case xlRangeAutoFormatColor3: XlRangeAutoFormatToString = "xlRangeAutoFormatColor3"
        Case xlRangeAutoFormatList1: XlRangeAutoFormatToString = "xlRangeAutoFormatList1"
        Case xlRangeAutoFormatList2: XlRangeAutoFormatToString = "xlRangeAutoFormatList2"
        Case xlRangeAutoFormatList3: XlRangeAutoFormatToString = "xlRangeAutoFormatList3"
        Case xlRangeAutoFormat3DEffects1: XlRangeAutoFormatToString = "xlRangeAutoFormat3DEffects1"
        Case xlRangeAutoFormat3DEffects2: XlRangeAutoFormatToString = "xlRangeAutoFormat3DEffects2"
        Case xlRangeAutoFormatLocalFormat1: XlRangeAutoFormatToString = "xlRangeAutoFormatLocalFormat1"
        Case xlRangeAutoFormatLocalFormat2: XlRangeAutoFormatToString = "xlRangeAutoFormatLocalFormat2"
        Case xlRangeAutoFormatAccounting4: XlRangeAutoFormatToString = "xlRangeAutoFormatAccounting4"
        Case xlRangeAutoFormatLocalFormat3: XlRangeAutoFormatToString = "xlRangeAutoFormatLocalFormat3"
        Case xlRangeAutoFormatLocalFormat4: XlRangeAutoFormatToString = "xlRangeAutoFormatLocalFormat4"
        Case xlRangeAutoFormatReport1: XlRangeAutoFormatToString = "xlRangeAutoFormatReport1"
        Case xlRangeAutoFormatReport2: XlRangeAutoFormatToString = "xlRangeAutoFormatReport2"
        Case xlRangeAutoFormatReport3: XlRangeAutoFormatToString = "xlRangeAutoFormatReport3"
        Case xlRangeAutoFormatReport4: XlRangeAutoFormatToString = "xlRangeAutoFormatReport4"
        Case xlRangeAutoFormatReport5: XlRangeAutoFormatToString = "xlRangeAutoFormatReport5"
        Case xlRangeAutoFormatReport6: XlRangeAutoFormatToString = "xlRangeAutoFormatReport6"
        Case xlRangeAutoFormatReport7: XlRangeAutoFormatToString = "xlRangeAutoFormatReport7"
        Case xlRangeAutoFormatReport8: XlRangeAutoFormatToString = "xlRangeAutoFormatReport8"
        Case xlRangeAutoFormatReport9: XlRangeAutoFormatToString = "xlRangeAutoFormatReport9"
        Case xlRangeAutoFormatReport10: XlRangeAutoFormatToString = "xlRangeAutoFormatReport10"
        Case xlRangeAutoFormatClassicPivotTable: XlRangeAutoFormatToString = "xlRangeAutoFormatClassicPivotTable"
        Case xlRangeAutoFormatTable1: XlRangeAutoFormatToString = "xlRangeAutoFormatTable1"
        Case xlRangeAutoFormatTable2: XlRangeAutoFormatToString = "xlRangeAutoFormatTable2"
        Case xlRangeAutoFormatTable3: XlRangeAutoFormatToString = "xlRangeAutoFormatTable3"
        Case xlRangeAutoFormatTable4: XlRangeAutoFormatToString = "xlRangeAutoFormatTable4"
        Case xlRangeAutoFormatTable5: XlRangeAutoFormatToString = "xlRangeAutoFormatTable5"
        Case xlRangeAutoFormatTable6: XlRangeAutoFormatToString = "xlRangeAutoFormatTable6"
        Case xlRangeAutoFormatTable7: XlRangeAutoFormatToString = "xlRangeAutoFormatTable7"
        Case xlRangeAutoFormatTable8: XlRangeAutoFormatToString = "xlRangeAutoFormatTable8"
        Case xlRangeAutoFormatTable9: XlRangeAutoFormatToString = "xlRangeAutoFormatTable9"
        Case xlRangeAutoFormatTable10: XlRangeAutoFormatToString = "xlRangeAutoFormatTable10"
        Case xlRangeAutoFormatPTNone: XlRangeAutoFormatToString = "xlRangeAutoFormatPTNone"
        Case xlRangeAutoFormatSimple: XlRangeAutoFormatToString = "xlRangeAutoFormatSimple"
        Case xlRangeAutoFormatNone: XlRangeAutoFormatToString = "xlRangeAutoFormatNone"
    End Select
End Function
