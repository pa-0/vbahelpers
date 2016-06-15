Attribute VB_Name = "wXlCalculatedMemberType"
Function XlCalculatedMemberTypeFromString(value As String) As XlCalculatedMemberType
    If IsNumeric(value) Then
        XlCalculatedMemberTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCalculatedMember": XlCalculatedMemberTypeFromString = xlCalculatedMember
        Case "xlCalculatedSet": XlCalculatedMemberTypeFromString = xlCalculatedSet
    End Select
End Function

Function XlCalculatedMemberTypeToString(value As XlCalculatedMemberType) As String
    Select Case value
        Case xlCalculatedMember: XlCalculatedMemberTypeToString = "xlCalculatedMember"
        Case xlCalculatedSet: XlCalculatedMemberTypeToString = "xlCalculatedSet"
    End Select
End Function
