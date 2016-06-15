Attribute VB_Name = "wMsoModeType"
Function MsoModeTypeFromString(value As String) As MsoModeType
    If IsNumeric(value) Then
        MsoModeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoModeModal": MsoModeTypeFromString = msoModeModal
        Case "msoModeAutoDown": MsoModeTypeFromString = msoModeAutoDown
        Case "msoModeModeless": MsoModeTypeFromString = msoModeModeless
    End Select
End Function

Function MsoModeTypeToString(value As MsoModeType) As String
    Select Case value
        Case msoModeModal: MsoModeTypeToString = "msoModeModal"
        Case msoModeAutoDown: MsoModeTypeToString = "msoModeAutoDown"
        Case msoModeModeless: MsoModeTypeToString = "msoModeModeless"
    End Select
End Function
