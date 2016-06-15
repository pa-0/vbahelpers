Attribute VB_Name = "wMsoFeatureInstall"
Function MsoFeatureInstallFromString(value As String) As MsoFeatureInstall
    If IsNumeric(value) Then
        MsoFeatureInstallFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFeatureInstallNone": MsoFeatureInstallFromString = msoFeatureInstallNone
        Case "msoFeatureInstallOnDemand": MsoFeatureInstallFromString = msoFeatureInstallOnDemand
        Case "msoFeatureInstallOnDemandWithUI": MsoFeatureInstallFromString = msoFeatureInstallOnDemandWithUI
    End Select
End Function

Function MsoFeatureInstallToString(value As MsoFeatureInstall) As String
    Select Case value
        Case msoFeatureInstallNone: MsoFeatureInstallToString = "msoFeatureInstallNone"
        Case msoFeatureInstallOnDemand: MsoFeatureInstallToString = "msoFeatureInstallOnDemand"
        Case msoFeatureInstallOnDemandWithUI: MsoFeatureInstallToString = "msoFeatureInstallOnDemandWithUI"
    End Select
End Function
