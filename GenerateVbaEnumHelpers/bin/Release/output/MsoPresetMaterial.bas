Attribute VB_Name = "wMsoPresetMaterial"
Function MsoPresetMaterialFromString(value As String) As MsoPresetMaterial
    If IsNumeric(value) Then
        MsoPresetMaterialFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoMaterialMatte": MsoPresetMaterialFromString = msoMaterialMatte
        Case "msoMaterialPlastic": MsoPresetMaterialFromString = msoMaterialPlastic
        Case "msoMaterialMetal": MsoPresetMaterialFromString = msoMaterialMetal
        Case "msoMaterialWireFrame": MsoPresetMaterialFromString = msoMaterialWireFrame
        Case "msoMaterialMatte2": MsoPresetMaterialFromString = msoMaterialMatte2
        Case "msoMaterialPlastic2": MsoPresetMaterialFromString = msoMaterialPlastic2
        Case "msoMaterialMetal2": MsoPresetMaterialFromString = msoMaterialMetal2
        Case "msoMaterialWarmMatte": MsoPresetMaterialFromString = msoMaterialWarmMatte
        Case "msoMaterialTranslucentPowder": MsoPresetMaterialFromString = msoMaterialTranslucentPowder
        Case "msoMaterialPowder": MsoPresetMaterialFromString = msoMaterialPowder
        Case "msoMaterialDarkEdge": MsoPresetMaterialFromString = msoMaterialDarkEdge
        Case "msoMaterialSoftEdge": MsoPresetMaterialFromString = msoMaterialSoftEdge
        Case "msoMaterialClear": MsoPresetMaterialFromString = msoMaterialClear
        Case "msoMaterialFlat": MsoPresetMaterialFromString = msoMaterialFlat
        Case "msoMaterialSoftMetal": MsoPresetMaterialFromString = msoMaterialSoftMetal
        Case "msoPresetMaterialMixed": MsoPresetMaterialFromString = msoPresetMaterialMixed
    End Select
End Function

Function MsoPresetMaterialToString(value As MsoPresetMaterial) As String
    Select Case value
        Case msoMaterialMatte: MsoPresetMaterialToString = "msoMaterialMatte"
        Case msoMaterialPlastic: MsoPresetMaterialToString = "msoMaterialPlastic"
        Case msoMaterialMetal: MsoPresetMaterialToString = "msoMaterialMetal"
        Case msoMaterialWireFrame: MsoPresetMaterialToString = "msoMaterialWireFrame"
        Case msoMaterialMatte2: MsoPresetMaterialToString = "msoMaterialMatte2"
        Case msoMaterialPlastic2: MsoPresetMaterialToString = "msoMaterialPlastic2"
        Case msoMaterialMetal2: MsoPresetMaterialToString = "msoMaterialMetal2"
        Case msoMaterialWarmMatte: MsoPresetMaterialToString = "msoMaterialWarmMatte"
        Case msoMaterialTranslucentPowder: MsoPresetMaterialToString = "msoMaterialTranslucentPowder"
        Case msoMaterialPowder: MsoPresetMaterialToString = "msoMaterialPowder"
        Case msoMaterialDarkEdge: MsoPresetMaterialToString = "msoMaterialDarkEdge"
        Case msoMaterialSoftEdge: MsoPresetMaterialToString = "msoMaterialSoftEdge"
        Case msoMaterialClear: MsoPresetMaterialToString = "msoMaterialClear"
        Case msoMaterialFlat: MsoPresetMaterialToString = "msoMaterialFlat"
        Case msoMaterialSoftMetal: MsoPresetMaterialToString = "msoMaterialSoftMetal"
        Case msoPresetMaterialMixed: MsoPresetMaterialToString = "msoPresetMaterialMixed"
    End Select
End Function
