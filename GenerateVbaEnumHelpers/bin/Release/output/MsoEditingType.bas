Attribute VB_Name = "wMsoEditingType"
Function MsoEditingTypeFromString(value As String) As MsoEditingType
    If IsNumeric(value) Then
        MsoEditingTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoEditingAuto": MsoEditingTypeFromString = msoEditingAuto
        Case "msoEditingCorner": MsoEditingTypeFromString = msoEditingCorner
        Case "msoEditingSmooth": MsoEditingTypeFromString = msoEditingSmooth
        Case "msoEditingSymmetric": MsoEditingTypeFromString = msoEditingSymmetric
    End Select
End Function

Function MsoEditingTypeToString(value As MsoEditingType) As String
    Select Case value
        Case msoEditingAuto: MsoEditingTypeToString = "msoEditingAuto"
        Case msoEditingCorner: MsoEditingTypeToString = "msoEditingCorner"
        Case msoEditingSmooth: MsoEditingTypeToString = "msoEditingSmooth"
        Case msoEditingSymmetric: MsoEditingTypeToString = "msoEditingSymmetric"
    End Select
End Function
