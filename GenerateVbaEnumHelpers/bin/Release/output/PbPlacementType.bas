Attribute VB_Name = "wPbPlacementType"
Function PbPlacementTypeFromString(value As String) As PbPlacementType
    If IsNumeric(value) Then
        PbPlacementTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPlacementLeft": PbPlacementTypeFromString = pbPlacementLeft
        Case "pbPlacementRight": PbPlacementTypeFromString = pbPlacementRight
        Case "pbPlacementCenter": PbPlacementTypeFromString = pbPlacementCenter
    End Select
End Function

Function PbPlacementTypeToString(value As PbPlacementType) As String
    Select Case value
        Case pbPlacementLeft: PbPlacementTypeToString = "pbPlacementLeft"
        Case pbPlacementRight: PbPlacementTypeToString = "pbPlacementRight"
        Case pbPlacementCenter: PbPlacementTypeToString = "pbPlacementCenter"
    End Select
End Function
