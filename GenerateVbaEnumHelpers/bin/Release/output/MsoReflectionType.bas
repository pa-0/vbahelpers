Attribute VB_Name = "wMsoReflectionType"
Function MsoReflectionTypeFromString(value As String) As MsoReflectionType
    If IsNumeric(value) Then
        MsoReflectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoReflectionTypeNone": MsoReflectionTypeFromString = msoReflectionTypeNone
        Case "msoReflectionType1": MsoReflectionTypeFromString = msoReflectionType1
        Case "msoReflectionType2": MsoReflectionTypeFromString = msoReflectionType2
        Case "msoReflectionType3": MsoReflectionTypeFromString = msoReflectionType3
        Case "msoReflectionType4": MsoReflectionTypeFromString = msoReflectionType4
        Case "msoReflectionType5": MsoReflectionTypeFromString = msoReflectionType5
        Case "msoReflectionType6": MsoReflectionTypeFromString = msoReflectionType6
        Case "msoReflectionType7": MsoReflectionTypeFromString = msoReflectionType7
        Case "msoReflectionType8": MsoReflectionTypeFromString = msoReflectionType8
        Case "msoReflectionType9": MsoReflectionTypeFromString = msoReflectionType9
        Case "msoReflectionTypeMixed": MsoReflectionTypeFromString = msoReflectionTypeMixed
    End Select
End Function

Function MsoReflectionTypeToString(value As MsoReflectionType) As String
    Select Case value
        Case msoReflectionTypeNone: MsoReflectionTypeToString = "msoReflectionTypeNone"
        Case msoReflectionType1: MsoReflectionTypeToString = "msoReflectionType1"
        Case msoReflectionType2: MsoReflectionTypeToString = "msoReflectionType2"
        Case msoReflectionType3: MsoReflectionTypeToString = "msoReflectionType3"
        Case msoReflectionType4: MsoReflectionTypeToString = "msoReflectionType4"
        Case msoReflectionType5: MsoReflectionTypeToString = "msoReflectionType5"
        Case msoReflectionType6: MsoReflectionTypeToString = "msoReflectionType6"
        Case msoReflectionType7: MsoReflectionTypeToString = "msoReflectionType7"
        Case msoReflectionType8: MsoReflectionTypeToString = "msoReflectionType8"
        Case msoReflectionType9: MsoReflectionTypeToString = "msoReflectionType9"
        Case msoReflectionTypeMixed: MsoReflectionTypeToString = "msoReflectionTypeMixed"
    End Select
End Function
