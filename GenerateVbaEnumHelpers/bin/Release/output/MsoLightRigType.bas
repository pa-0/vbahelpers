Attribute VB_Name = "wMsoLightRigType"
Function MsoLightRigTypeFromString(value As String) As MsoLightRigType
    If IsNumeric(value) Then
        MsoLightRigTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLightRigLegacyFlat1": MsoLightRigTypeFromString = msoLightRigLegacyFlat1
        Case "msoLightRigLegacyFlat2": MsoLightRigTypeFromString = msoLightRigLegacyFlat2
        Case "msoLightRigLegacyFlat3": MsoLightRigTypeFromString = msoLightRigLegacyFlat3
        Case "msoLightRigLegacyFlat4": MsoLightRigTypeFromString = msoLightRigLegacyFlat4
        Case "msoLightRigLegacyNormal1": MsoLightRigTypeFromString = msoLightRigLegacyNormal1
        Case "msoLightRigLegacyNormal2": MsoLightRigTypeFromString = msoLightRigLegacyNormal2
        Case "msoLightRigLegacyNormal3": MsoLightRigTypeFromString = msoLightRigLegacyNormal3
        Case "msoLightRigLegacyNormal4": MsoLightRigTypeFromString = msoLightRigLegacyNormal4
        Case "msoLightRigLegacyHarsh1": MsoLightRigTypeFromString = msoLightRigLegacyHarsh1
        Case "msoLightRigLegacyHarsh2": MsoLightRigTypeFromString = msoLightRigLegacyHarsh2
        Case "msoLightRigLegacyHarsh3": MsoLightRigTypeFromString = msoLightRigLegacyHarsh3
        Case "msoLightRigLegacyHarsh4": MsoLightRigTypeFromString = msoLightRigLegacyHarsh4
        Case "msoLightRigThreePoint": MsoLightRigTypeFromString = msoLightRigThreePoint
        Case "msoLightRigBalanced": MsoLightRigTypeFromString = msoLightRigBalanced
        Case "msoLightRigSoft": MsoLightRigTypeFromString = msoLightRigSoft
        Case "msoLightRigHarsh": MsoLightRigTypeFromString = msoLightRigHarsh
        Case "msoLightRigFlood": MsoLightRigTypeFromString = msoLightRigFlood
        Case "msoLightRigContrasting": MsoLightRigTypeFromString = msoLightRigContrasting
        Case "msoLightRigMorning": MsoLightRigTypeFromString = msoLightRigMorning
        Case "msoLightRigSunrise": MsoLightRigTypeFromString = msoLightRigSunrise
        Case "msoLightRigSunset": MsoLightRigTypeFromString = msoLightRigSunset
        Case "msoLightRigChilly": MsoLightRigTypeFromString = msoLightRigChilly
        Case "msoLightRigFreezing": MsoLightRigTypeFromString = msoLightRigFreezing
        Case "msoLightRigFlat": MsoLightRigTypeFromString = msoLightRigFlat
        Case "msoLightRigTwoPoint": MsoLightRigTypeFromString = msoLightRigTwoPoint
        Case "msoLightRigGlow": MsoLightRigTypeFromString = msoLightRigGlow
        Case "msoLightRigBrightRoom": MsoLightRigTypeFromString = msoLightRigBrightRoom
        Case "msoLightRigMixed": MsoLightRigTypeFromString = msoLightRigMixed
    End Select
End Function

Function MsoLightRigTypeToString(value As MsoLightRigType) As String
    Select Case value
        Case msoLightRigLegacyFlat1: MsoLightRigTypeToString = "msoLightRigLegacyFlat1"
        Case msoLightRigLegacyFlat2: MsoLightRigTypeToString = "msoLightRigLegacyFlat2"
        Case msoLightRigLegacyFlat3: MsoLightRigTypeToString = "msoLightRigLegacyFlat3"
        Case msoLightRigLegacyFlat4: MsoLightRigTypeToString = "msoLightRigLegacyFlat4"
        Case msoLightRigLegacyNormal1: MsoLightRigTypeToString = "msoLightRigLegacyNormal1"
        Case msoLightRigLegacyNormal2: MsoLightRigTypeToString = "msoLightRigLegacyNormal2"
        Case msoLightRigLegacyNormal3: MsoLightRigTypeToString = "msoLightRigLegacyNormal3"
        Case msoLightRigLegacyNormal4: MsoLightRigTypeToString = "msoLightRigLegacyNormal4"
        Case msoLightRigLegacyHarsh1: MsoLightRigTypeToString = "msoLightRigLegacyHarsh1"
        Case msoLightRigLegacyHarsh2: MsoLightRigTypeToString = "msoLightRigLegacyHarsh2"
        Case msoLightRigLegacyHarsh3: MsoLightRigTypeToString = "msoLightRigLegacyHarsh3"
        Case msoLightRigLegacyHarsh4: MsoLightRigTypeToString = "msoLightRigLegacyHarsh4"
        Case msoLightRigThreePoint: MsoLightRigTypeToString = "msoLightRigThreePoint"
        Case msoLightRigBalanced: MsoLightRigTypeToString = "msoLightRigBalanced"
        Case msoLightRigSoft: MsoLightRigTypeToString = "msoLightRigSoft"
        Case msoLightRigHarsh: MsoLightRigTypeToString = "msoLightRigHarsh"
        Case msoLightRigFlood: MsoLightRigTypeToString = "msoLightRigFlood"
        Case msoLightRigContrasting: MsoLightRigTypeToString = "msoLightRigContrasting"
        Case msoLightRigMorning: MsoLightRigTypeToString = "msoLightRigMorning"
        Case msoLightRigSunrise: MsoLightRigTypeToString = "msoLightRigSunrise"
        Case msoLightRigSunset: MsoLightRigTypeToString = "msoLightRigSunset"
        Case msoLightRigChilly: MsoLightRigTypeToString = "msoLightRigChilly"
        Case msoLightRigFreezing: MsoLightRigTypeToString = "msoLightRigFreezing"
        Case msoLightRigFlat: MsoLightRigTypeToString = "msoLightRigFlat"
        Case msoLightRigTwoPoint: MsoLightRigTypeToString = "msoLightRigTwoPoint"
        Case msoLightRigGlow: MsoLightRigTypeToString = "msoLightRigGlow"
        Case msoLightRigBrightRoom: MsoLightRigTypeToString = "msoLightRigBrightRoom"
        Case msoLightRigMixed: MsoLightRigTypeToString = "msoLightRigMixed"
    End Select
End Function
