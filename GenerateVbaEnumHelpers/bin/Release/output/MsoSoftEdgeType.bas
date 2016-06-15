Attribute VB_Name = "wMsoSoftEdgeType"
Function MsoSoftEdgeTypeFromString(value As String) As MsoSoftEdgeType
    If IsNumeric(value) Then
        MsoSoftEdgeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSoftEdgeTypeNone": MsoSoftEdgeTypeFromString = msoSoftEdgeTypeNone
        Case "msoSoftEdgeType1": MsoSoftEdgeTypeFromString = msoSoftEdgeType1
        Case "msoSoftEdgeType2": MsoSoftEdgeTypeFromString = msoSoftEdgeType2
        Case "msoSoftEdgeType3": MsoSoftEdgeTypeFromString = msoSoftEdgeType3
        Case "msoSoftEdgeType4": MsoSoftEdgeTypeFromString = msoSoftEdgeType4
        Case "msoSoftEdgeType5": MsoSoftEdgeTypeFromString = msoSoftEdgeType5
        Case "msoSoftEdgeType6": MsoSoftEdgeTypeFromString = msoSoftEdgeType6
        Case "msoSoftEdgeTypeMixed": MsoSoftEdgeTypeFromString = msoSoftEdgeTypeMixed
    End Select
End Function

Function MsoSoftEdgeTypeToString(value As MsoSoftEdgeType) As String
    Select Case value
        Case msoSoftEdgeTypeNone: MsoSoftEdgeTypeToString = "msoSoftEdgeTypeNone"
        Case msoSoftEdgeType1: MsoSoftEdgeTypeToString = "msoSoftEdgeType1"
        Case msoSoftEdgeType2: MsoSoftEdgeTypeToString = "msoSoftEdgeType2"
        Case msoSoftEdgeType3: MsoSoftEdgeTypeToString = "msoSoftEdgeType3"
        Case msoSoftEdgeType4: MsoSoftEdgeTypeToString = "msoSoftEdgeType4"
        Case msoSoftEdgeType5: MsoSoftEdgeTypeToString = "msoSoftEdgeType5"
        Case msoSoftEdgeType6: MsoSoftEdgeTypeToString = "msoSoftEdgeType6"
        Case msoSoftEdgeTypeMixed: MsoSoftEdgeTypeToString = "msoSoftEdgeTypeMixed"
    End Select
End Function
