Attribute VB_Name = "wPbShowDialog"
Function PbShowDialogFromString(value As String) As PbShowDialog
    If IsNumeric(value) Then
        PbShowDialogFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbDefaultBehavior": PbShowDialogFromString = pbDefaultBehavior
        Case "PbShowDialog": PbShowDialogFromString = PbShowDialog
        Case "pbSuppressDialog": PbShowDialogFromString = pbSuppressDialog
    End Select
End Function

Function PbShowDialogToString(value As PbShowDialog) As String
    Select Case value
        Case pbDefaultBehavior: PbShowDialogToString = "pbDefaultBehavior"
        Case PbShowDialog: PbShowDialogToString = "PbShowDialog"
        Case pbSuppressDialog: PbShowDialogToString = "pbSuppressDialog"
    End Select
End Function
