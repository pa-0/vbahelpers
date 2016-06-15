Attribute VB_Name = "wXlToolbarProtection"
Function XlToolbarProtectionFromString(value As String) As XlToolbarProtection
    If IsNumeric(value) Then
        XlToolbarProtectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNoButtonChanges": XlToolbarProtectionFromString = xlNoButtonChanges
        Case "xlNoShapeChanges": XlToolbarProtectionFromString = xlNoShapeChanges
        Case "xlNoDockingChanges": XlToolbarProtectionFromString = xlNoDockingChanges
        Case "xlNoChanges": XlToolbarProtectionFromString = xlNoChanges
        Case "xlToolbarProtectionNone": XlToolbarProtectionFromString = xlToolbarProtectionNone
    End Select
End Function

Function XlToolbarProtectionToString(value As XlToolbarProtection) As String
    Select Case value
        Case xlNoButtonChanges: XlToolbarProtectionToString = "xlNoButtonChanges"
        Case xlNoShapeChanges: XlToolbarProtectionToString = "xlNoShapeChanges"
        Case xlNoDockingChanges: XlToolbarProtectionToString = "xlNoDockingChanges"
        Case xlNoChanges: XlToolbarProtectionToString = "xlNoChanges"
        Case xlToolbarProtectionNone: XlToolbarProtectionToString = "xlToolbarProtectionNone"
    End Select
End Function
