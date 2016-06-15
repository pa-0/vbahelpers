Attribute VB_Name = "wPbBuildingBlockType"
Function PbBuildingBlockTypeFromString(value As String) As PbBuildingBlockType
    If IsNumeric(value) Then
        PbBuildingBlockTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbBBNone": PbBuildingBlockTypeFromString = pbBBNone
        Case "pbBBBuiltIn": PbBuildingBlockTypeFromString = pbBBBuiltIn
        Case "pbBBDownloaded": PbBuildingBlockTypeFromString = pbBBDownloaded
        Case "pbBBUser": PbBuildingBlockTypeFromString = pbBBUser
        Case "pbBBWorkGroup": PbBuildingBlockTypeFromString = pbBBWorkGroup
    End Select
End Function

Function PbBuildingBlockTypeToString(value As PbBuildingBlockType) As String
    Select Case value
        Case pbBBNone: PbBuildingBlockTypeToString = "pbBBNone"
        Case pbBBBuiltIn: PbBuildingBlockTypeToString = "pbBBBuiltIn"
        Case pbBBDownloaded: PbBuildingBlockTypeToString = "pbBBDownloaded"
        Case pbBBUser: PbBuildingBlockTypeToString = "pbBBUser"
        Case pbBBWorkGroup: PbBuildingBlockTypeToString = "pbBBWorkGroup"
    End Select
End Function
