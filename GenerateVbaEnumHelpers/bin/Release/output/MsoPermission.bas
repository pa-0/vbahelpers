Attribute VB_Name = "wMsoPermission"
Function MsoPermissionFromString(value As String) As MsoPermission
    If IsNumeric(value) Then
        MsoPermissionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoPermissionView": MsoPermissionFromString = msoPermissionView
        Case "msoPermissionRead": MsoPermissionFromString = msoPermissionRead
        Case "msoPermissionEdit": MsoPermissionFromString = msoPermissionEdit
        Case "msoPermissionSave": MsoPermissionFromString = msoPermissionSave
        Case "msoPermissionExtract": MsoPermissionFromString = msoPermissionExtract
        Case "msoPermissionChange": MsoPermissionFromString = msoPermissionChange
        Case "msoPermissionPrint": MsoPermissionFromString = msoPermissionPrint
        Case "msoPermissionObjModel": MsoPermissionFromString = msoPermissionObjModel
        Case "msoPermissionFullControl": MsoPermissionFromString = msoPermissionFullControl
        Case "msoPermissionAllCommon": MsoPermissionFromString = msoPermissionAllCommon
    End Select
End Function

Function MsoPermissionToString(value As MsoPermission) As String
    Select Case value
        Case msoPermissionView: MsoPermissionToString = "msoPermissionView"
        Case msoPermissionRead: MsoPermissionToString = "msoPermissionRead"
        Case msoPermissionEdit: MsoPermissionToString = "msoPermissionEdit"
        Case msoPermissionSave: MsoPermissionToString = "msoPermissionSave"
        Case msoPermissionExtract: MsoPermissionToString = "msoPermissionExtract"
        Case msoPermissionChange: MsoPermissionToString = "msoPermissionChange"
        Case msoPermissionPrint: MsoPermissionToString = "msoPermissionPrint"
        Case msoPermissionObjModel: MsoPermissionToString = "msoPermissionObjModel"
        Case msoPermissionFullControl: MsoPermissionToString = "msoPermissionFullControl"
        Case msoPermissionAllCommon: MsoPermissionToString = "msoPermissionAllCommon"
    End Select
End Function
