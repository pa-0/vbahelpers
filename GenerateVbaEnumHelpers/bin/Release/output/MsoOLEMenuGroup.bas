Attribute VB_Name = "wMsoOLEMenuGroup"
Function MsoOLEMenuGroupFromString(value As String) As MsoOLEMenuGroup
    If IsNumeric(value) Then
        MsoOLEMenuGroupFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoOLEMenuGroupFile": MsoOLEMenuGroupFromString = msoOLEMenuGroupFile
        Case "msoOLEMenuGroupEdit": MsoOLEMenuGroupFromString = msoOLEMenuGroupEdit
        Case "msoOLEMenuGroupContainer": MsoOLEMenuGroupFromString = msoOLEMenuGroupContainer
        Case "msoOLEMenuGroupObject": MsoOLEMenuGroupFromString = msoOLEMenuGroupObject
        Case "msoOLEMenuGroupWindow": MsoOLEMenuGroupFromString = msoOLEMenuGroupWindow
        Case "msoOLEMenuGroupHelp": MsoOLEMenuGroupFromString = msoOLEMenuGroupHelp
        Case "msoOLEMenuGroupNone": MsoOLEMenuGroupFromString = msoOLEMenuGroupNone
    End Select
End Function

Function MsoOLEMenuGroupToString(value As MsoOLEMenuGroup) As String
    Select Case value
        Case msoOLEMenuGroupFile: MsoOLEMenuGroupToString = "msoOLEMenuGroupFile"
        Case msoOLEMenuGroupEdit: MsoOLEMenuGroupToString = "msoOLEMenuGroupEdit"
        Case msoOLEMenuGroupContainer: MsoOLEMenuGroupToString = "msoOLEMenuGroupContainer"
        Case msoOLEMenuGroupObject: MsoOLEMenuGroupToString = "msoOLEMenuGroupObject"
        Case msoOLEMenuGroupWindow: MsoOLEMenuGroupToString = "msoOLEMenuGroupWindow"
        Case msoOLEMenuGroupHelp: MsoOLEMenuGroupToString = "msoOLEMenuGroupHelp"
        Case msoOLEMenuGroupNone: MsoOLEMenuGroupToString = "msoOLEMenuGroupNone"
    End Select
End Function
