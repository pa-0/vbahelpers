Attribute VB_Name = "wWdTableFormat"
Function WdTableFormatFromString(value As String) As WdTableFormat
    If IsNumeric(value) Then
        WdTableFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTableFormatNone": WdTableFormatFromString = wdTableFormatNone
        Case "wdTableFormatSimple1": WdTableFormatFromString = wdTableFormatSimple1
        Case "wdTableFormatSimple2": WdTableFormatFromString = wdTableFormatSimple2
        Case "wdTableFormatSimple3": WdTableFormatFromString = wdTableFormatSimple3
        Case "wdTableFormatClassic1": WdTableFormatFromString = wdTableFormatClassic1
        Case "wdTableFormatClassic2": WdTableFormatFromString = wdTableFormatClassic2
        Case "wdTableFormatClassic3": WdTableFormatFromString = wdTableFormatClassic3
        Case "wdTableFormatClassic4": WdTableFormatFromString = wdTableFormatClassic4
        Case "wdTableFormatColorful1": WdTableFormatFromString = wdTableFormatColorful1
        Case "wdTableFormatColorful2": WdTableFormatFromString = wdTableFormatColorful2
        Case "wdTableFormatColorful3": WdTableFormatFromString = wdTableFormatColorful3
        Case "wdTableFormatColumns1": WdTableFormatFromString = wdTableFormatColumns1
        Case "wdTableFormatColumns2": WdTableFormatFromString = wdTableFormatColumns2
        Case "wdTableFormatColumns3": WdTableFormatFromString = wdTableFormatColumns3
        Case "wdTableFormatColumns4": WdTableFormatFromString = wdTableFormatColumns4
        Case "wdTableFormatColumns5": WdTableFormatFromString = wdTableFormatColumns5
        Case "wdTableFormatGrid1": WdTableFormatFromString = wdTableFormatGrid1
        Case "wdTableFormatGrid2": WdTableFormatFromString = wdTableFormatGrid2
        Case "wdTableFormatGrid3": WdTableFormatFromString = wdTableFormatGrid3
        Case "wdTableFormatGrid4": WdTableFormatFromString = wdTableFormatGrid4
        Case "wdTableFormatGrid5": WdTableFormatFromString = wdTableFormatGrid5
        Case "wdTableFormatGrid6": WdTableFormatFromString = wdTableFormatGrid6
        Case "wdTableFormatGrid7": WdTableFormatFromString = wdTableFormatGrid7
        Case "wdTableFormatGrid8": WdTableFormatFromString = wdTableFormatGrid8
        Case "wdTableFormatList1": WdTableFormatFromString = wdTableFormatList1
        Case "wdTableFormatList2": WdTableFormatFromString = wdTableFormatList2
        Case "wdTableFormatList3": WdTableFormatFromString = wdTableFormatList3
        Case "wdTableFormatList4": WdTableFormatFromString = wdTableFormatList4
        Case "wdTableFormatList5": WdTableFormatFromString = wdTableFormatList5
        Case "wdTableFormatList6": WdTableFormatFromString = wdTableFormatList6
        Case "wdTableFormatList7": WdTableFormatFromString = wdTableFormatList7
        Case "wdTableFormatList8": WdTableFormatFromString = wdTableFormatList8
        Case "wdTableFormat3DEffects1": WdTableFormatFromString = wdTableFormat3DEffects1
        Case "wdTableFormat3DEffects2": WdTableFormatFromString = wdTableFormat3DEffects2
        Case "wdTableFormat3DEffects3": WdTableFormatFromString = wdTableFormat3DEffects3
        Case "wdTableFormatContemporary": WdTableFormatFromString = wdTableFormatContemporary
        Case "wdTableFormatElegant": WdTableFormatFromString = wdTableFormatElegant
        Case "wdTableFormatProfessional": WdTableFormatFromString = wdTableFormatProfessional
        Case "wdTableFormatSubtle1": WdTableFormatFromString = wdTableFormatSubtle1
        Case "wdTableFormatSubtle2": WdTableFormatFromString = wdTableFormatSubtle2
        Case "wdTableFormatWeb1": WdTableFormatFromString = wdTableFormatWeb1
        Case "wdTableFormatWeb2": WdTableFormatFromString = wdTableFormatWeb2
        Case "wdTableFormatWeb3": WdTableFormatFromString = wdTableFormatWeb3
    End Select
End Function

Function WdTableFormatToString(value As WdTableFormat) As String
    Select Case value
        Case wdTableFormatNone: WdTableFormatToString = "wdTableFormatNone"
        Case wdTableFormatSimple1: WdTableFormatToString = "wdTableFormatSimple1"
        Case wdTableFormatSimple2: WdTableFormatToString = "wdTableFormatSimple2"
        Case wdTableFormatSimple3: WdTableFormatToString = "wdTableFormatSimple3"
        Case wdTableFormatClassic1: WdTableFormatToString = "wdTableFormatClassic1"
        Case wdTableFormatClassic2: WdTableFormatToString = "wdTableFormatClassic2"
        Case wdTableFormatClassic3: WdTableFormatToString = "wdTableFormatClassic3"
        Case wdTableFormatClassic4: WdTableFormatToString = "wdTableFormatClassic4"
        Case wdTableFormatColorful1: WdTableFormatToString = "wdTableFormatColorful1"
        Case wdTableFormatColorful2: WdTableFormatToString = "wdTableFormatColorful2"
        Case wdTableFormatColorful3: WdTableFormatToString = "wdTableFormatColorful3"
        Case wdTableFormatColumns1: WdTableFormatToString = "wdTableFormatColumns1"
        Case wdTableFormatColumns2: WdTableFormatToString = "wdTableFormatColumns2"
        Case wdTableFormatColumns3: WdTableFormatToString = "wdTableFormatColumns3"
        Case wdTableFormatColumns4: WdTableFormatToString = "wdTableFormatColumns4"
        Case wdTableFormatColumns5: WdTableFormatToString = "wdTableFormatColumns5"
        Case wdTableFormatGrid1: WdTableFormatToString = "wdTableFormatGrid1"
        Case wdTableFormatGrid2: WdTableFormatToString = "wdTableFormatGrid2"
        Case wdTableFormatGrid3: WdTableFormatToString = "wdTableFormatGrid3"
        Case wdTableFormatGrid4: WdTableFormatToString = "wdTableFormatGrid4"
        Case wdTableFormatGrid5: WdTableFormatToString = "wdTableFormatGrid5"
        Case wdTableFormatGrid6: WdTableFormatToString = "wdTableFormatGrid6"
        Case wdTableFormatGrid7: WdTableFormatToString = "wdTableFormatGrid7"
        Case wdTableFormatGrid8: WdTableFormatToString = "wdTableFormatGrid8"
        Case wdTableFormatList1: WdTableFormatToString = "wdTableFormatList1"
        Case wdTableFormatList2: WdTableFormatToString = "wdTableFormatList2"
        Case wdTableFormatList3: WdTableFormatToString = "wdTableFormatList3"
        Case wdTableFormatList4: WdTableFormatToString = "wdTableFormatList4"
        Case wdTableFormatList5: WdTableFormatToString = "wdTableFormatList5"
        Case wdTableFormatList6: WdTableFormatToString = "wdTableFormatList6"
        Case wdTableFormatList7: WdTableFormatToString = "wdTableFormatList7"
        Case wdTableFormatList8: WdTableFormatToString = "wdTableFormatList8"
        Case wdTableFormat3DEffects1: WdTableFormatToString = "wdTableFormat3DEffects1"
        Case wdTableFormat3DEffects2: WdTableFormatToString = "wdTableFormat3DEffects2"
        Case wdTableFormat3DEffects3: WdTableFormatToString = "wdTableFormat3DEffects3"
        Case wdTableFormatContemporary: WdTableFormatToString = "wdTableFormatContemporary"
        Case wdTableFormatElegant: WdTableFormatToString = "wdTableFormatElegant"
        Case wdTableFormatProfessional: WdTableFormatToString = "wdTableFormatProfessional"
        Case wdTableFormatSubtle1: WdTableFormatToString = "wdTableFormatSubtle1"
        Case wdTableFormatSubtle2: WdTableFormatToString = "wdTableFormatSubtle2"
        Case wdTableFormatWeb1: WdTableFormatToString = "wdTableFormatWeb1"
        Case wdTableFormatWeb2: WdTableFormatToString = "wdTableFormatWeb2"
        Case wdTableFormatWeb3: WdTableFormatToString = "wdTableFormatWeb3"
    End Select
End Function
