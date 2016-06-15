Attribute VB_Name = "wMsoAnimType"
Function MsoAnimTypeFromString(value As String) As MsoAnimType
    If IsNumeric(value) Then
        MsoAnimTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimTypeNone": MsoAnimTypeFromString = msoAnimTypeNone
        Case "msoAnimTypeMotion": MsoAnimTypeFromString = msoAnimTypeMotion
        Case "msoAnimTypeColor": MsoAnimTypeFromString = msoAnimTypeColor
        Case "msoAnimTypeScale": MsoAnimTypeFromString = msoAnimTypeScale
        Case "msoAnimTypeRotation": MsoAnimTypeFromString = msoAnimTypeRotation
        Case "msoAnimTypeProperty": MsoAnimTypeFromString = msoAnimTypeProperty
        Case "msoAnimTypeCommand": MsoAnimTypeFromString = msoAnimTypeCommand
        Case "msoAnimTypeFilter": MsoAnimTypeFromString = msoAnimTypeFilter
        Case "msoAnimTypeSet": MsoAnimTypeFromString = msoAnimTypeSet
        Case "msoAnimTypeMixed": MsoAnimTypeFromString = msoAnimTypeMixed
    End Select
End Function

Function MsoAnimTypeToString(value As MsoAnimType) As String
    Select Case value
        Case msoAnimTypeNone: MsoAnimTypeToString = "msoAnimTypeNone"
        Case msoAnimTypeMotion: MsoAnimTypeToString = "msoAnimTypeMotion"
        Case msoAnimTypeColor: MsoAnimTypeToString = "msoAnimTypeColor"
        Case msoAnimTypeScale: MsoAnimTypeToString = "msoAnimTypeScale"
        Case msoAnimTypeRotation: MsoAnimTypeToString = "msoAnimTypeRotation"
        Case msoAnimTypeProperty: MsoAnimTypeToString = "msoAnimTypeProperty"
        Case msoAnimTypeCommand: MsoAnimTypeToString = "msoAnimTypeCommand"
        Case msoAnimTypeFilter: MsoAnimTypeToString = "msoAnimTypeFilter"
        Case msoAnimTypeSet: MsoAnimTypeToString = "msoAnimTypeSet"
        Case msoAnimTypeMixed: MsoAnimTypeToString = "msoAnimTypeMixed"
    End Select
End Function
