Attribute VB_Name = "wMsoAlignCmd"
Function MsoAlignCmdFromString(value As String) As MsoAlignCmd
    If IsNumeric(value) Then
        MsoAlignCmdFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAlignLefts": MsoAlignCmdFromString = msoAlignLefts
        Case "msoAlignCenters": MsoAlignCmdFromString = msoAlignCenters
        Case "msoAlignRights": MsoAlignCmdFromString = msoAlignRights
        Case "msoAlignTops": MsoAlignCmdFromString = msoAlignTops
        Case "msoAlignMiddles": MsoAlignCmdFromString = msoAlignMiddles
        Case "msoAlignBottoms": MsoAlignCmdFromString = msoAlignBottoms
    End Select
End Function

Function MsoAlignCmdToString(value As MsoAlignCmd) As String
    Select Case value
        Case msoAlignLefts: MsoAlignCmdToString = "msoAlignLefts"
        Case msoAlignCenters: MsoAlignCmdToString = "msoAlignCenters"
        Case msoAlignRights: MsoAlignCmdToString = "msoAlignRights"
        Case msoAlignTops: MsoAlignCmdToString = "msoAlignTops"
        Case msoAlignMiddles: MsoAlignCmdToString = "msoAlignMiddles"
        Case msoAlignBottoms: MsoAlignCmdToString = "msoAlignBottoms"
    End Select
End Function
