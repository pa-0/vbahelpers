Attribute VB_Name = "wMsoIodGroup"
Function MsoIodGroupFromString(value As String) As MsoIodGroup
    If IsNumeric(value) Then
        MsoIodGroupFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoIodGroupPIAs": MsoIodGroupFromString = msoIodGroupPIAs
        Case "msoIodGroupVSTOR35Mgd": MsoIodGroupFromString = msoIodGroupVSTOR35Mgd
        Case "msoIodGroupVSTOR40Mgd": MsoIodGroupFromString = msoIodGroupVSTOR40Mgd
    End Select
End Function

Function MsoIodGroupToString(value As MsoIodGroup) As String
    Select Case value
        Case msoIodGroupPIAs: MsoIodGroupToString = "msoIodGroupPIAs"
        Case msoIodGroupVSTOR35Mgd: MsoIodGroupToString = "msoIodGroupVSTOR35Mgd"
        Case msoIodGroupVSTOR40Mgd: MsoIodGroupToString = "msoIodGroupVSTOR40Mgd"
    End Select
End Function
