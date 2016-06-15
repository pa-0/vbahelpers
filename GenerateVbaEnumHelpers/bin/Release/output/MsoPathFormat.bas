Attribute VB_Name = "wMsoPathFormat"
Function MsoPathFormatFromString(value As String) As MsoPathFormat
    If IsNumeric(value) Then
        MsoPathFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoPathTypeNone": MsoPathFormatFromString = msoPathTypeNone
        Case "msoPathType1": MsoPathFormatFromString = msoPathType1
        Case "msoPathType2": MsoPathFormatFromString = msoPathType2
        Case "msoPathType3": MsoPathFormatFromString = msoPathType3
        Case "msoPathType4": MsoPathFormatFromString = msoPathType4
        Case "msoPathTypeMixed": MsoPathFormatFromString = msoPathTypeMixed
    End Select
End Function

Function MsoPathFormatToString(value As MsoPathFormat) As String
    Select Case value
        Case msoPathTypeNone: MsoPathFormatToString = "msoPathTypeNone"
        Case msoPathType1: MsoPathFormatToString = "msoPathType1"
        Case msoPathType2: MsoPathFormatToString = "msoPathType2"
        Case msoPathType3: MsoPathFormatToString = "msoPathType3"
        Case msoPathType4: MsoPathFormatToString = "msoPathType4"
        Case msoPathTypeMixed: MsoPathFormatToString = "msoPathTypeMixed"
    End Select
End Function
