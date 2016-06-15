Attribute VB_Name = "wMsoBarType"
Function MsoBarTypeFromString(value As String) As MsoBarType
    If IsNumeric(value) Then
        MsoBarTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBarTypeNormal": MsoBarTypeFromString = msoBarTypeNormal
        Case "msoBarTypeMenuBar": MsoBarTypeFromString = msoBarTypeMenuBar
        Case "msoBarTypePopup": MsoBarTypeFromString = msoBarTypePopup
    End Select
End Function

Function MsoBarTypeToString(value As MsoBarType) As String
    Select Case value
        Case msoBarTypeNormal: MsoBarTypeToString = "msoBarTypeNormal"
        Case msoBarTypeMenuBar: MsoBarTypeToString = "msoBarTypeMenuBar"
        Case msoBarTypePopup: MsoBarTypeToString = "msoBarTypePopup"
    End Select
End Function
