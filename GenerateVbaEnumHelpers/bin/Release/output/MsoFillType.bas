Attribute VB_Name = "wMsoFillType"
Function MsoFillTypeFromString(value As String) As MsoFillType
    If IsNumeric(value) Then
        MsoFillTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFillSolid": MsoFillTypeFromString = msoFillSolid
        Case "msoFillPatterned": MsoFillTypeFromString = msoFillPatterned
        Case "msoFillGradient": MsoFillTypeFromString = msoFillGradient
        Case "msoFillTextured": MsoFillTypeFromString = msoFillTextured
        Case "msoFillBackground": MsoFillTypeFromString = msoFillBackground
        Case "msoFillPicture": MsoFillTypeFromString = msoFillPicture
        Case "msoFillMixed": MsoFillTypeFromString = msoFillMixed
    End Select
End Function

Function MsoFillTypeToString(value As MsoFillType) As String
    Select Case value
        Case msoFillSolid: MsoFillTypeToString = "msoFillSolid"
        Case msoFillPatterned: MsoFillTypeToString = "msoFillPatterned"
        Case msoFillGradient: MsoFillTypeToString = "msoFillGradient"
        Case msoFillTextured: MsoFillTypeToString = "msoFillTextured"
        Case msoFillBackground: MsoFillTypeToString = "msoFillBackground"
        Case msoFillPicture: MsoFillTypeToString = "msoFillPicture"
        Case msoFillMixed: MsoFillTypeToString = "msoFillMixed"
    End Select
End Function
