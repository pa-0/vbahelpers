Attribute VB_Name = "wMsoBarProtection"
Function MsoBarProtectionFromString(value As String) As MsoBarProtection
    If IsNumeric(value) Then
        MsoBarProtectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBarNoProtection": MsoBarProtectionFromString = msoBarNoProtection
        Case "msoBarNoCustomize": MsoBarProtectionFromString = msoBarNoCustomize
        Case "msoBarNoResize": MsoBarProtectionFromString = msoBarNoResize
        Case "msoBarNoMove": MsoBarProtectionFromString = msoBarNoMove
        Case "msoBarNoChangeVisible": MsoBarProtectionFromString = msoBarNoChangeVisible
        Case "msoBarNoChangeDock": MsoBarProtectionFromString = msoBarNoChangeDock
        Case "msoBarNoVerticalDock": MsoBarProtectionFromString = msoBarNoVerticalDock
        Case "msoBarNoHorizontalDock": MsoBarProtectionFromString = msoBarNoHorizontalDock
    End Select
End Function

Function MsoBarProtectionToString(value As MsoBarProtection) As String
    Select Case value
        Case msoBarNoProtection: MsoBarProtectionToString = "msoBarNoProtection"
        Case msoBarNoCustomize: MsoBarProtectionToString = "msoBarNoCustomize"
        Case msoBarNoResize: MsoBarProtectionToString = "msoBarNoResize"
        Case msoBarNoMove: MsoBarProtectionToString = "msoBarNoMove"
        Case msoBarNoChangeVisible: MsoBarProtectionToString = "msoBarNoChangeVisible"
        Case msoBarNoChangeDock: MsoBarProtectionToString = "msoBarNoChangeDock"
        Case msoBarNoVerticalDock: MsoBarProtectionToString = "msoBarNoVerticalDock"
        Case msoBarNoHorizontalDock: MsoBarProtectionToString = "msoBarNoHorizontalDock"
    End Select
End Function
