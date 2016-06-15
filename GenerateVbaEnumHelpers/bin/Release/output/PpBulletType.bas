Attribute VB_Name = "wPpBulletType"
Function PpBulletTypeFromString(value As String) As PpBulletType
    If IsNumeric(value) Then
        PpBulletTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppBulletNone": PpBulletTypeFromString = ppBulletNone
        Case "ppBulletUnnumbered": PpBulletTypeFromString = ppBulletUnnumbered
        Case "ppBulletNumbered": PpBulletTypeFromString = ppBulletNumbered
        Case "ppBulletPicture": PpBulletTypeFromString = ppBulletPicture
        Case "ppBulletMixed": PpBulletTypeFromString = ppBulletMixed
    End Select
End Function

Function PpBulletTypeToString(value As PpBulletType) As String
    Select Case value
        Case ppBulletNone: PpBulletTypeToString = "ppBulletNone"
        Case ppBulletUnnumbered: PpBulletTypeToString = "ppBulletUnnumbered"
        Case ppBulletNumbered: PpBulletTypeToString = "ppBulletNumbered"
        Case ppBulletPicture: PpBulletTypeToString = "ppBulletPicture"
        Case ppBulletMixed: PpBulletTypeToString = "ppBulletMixed"
    End Select
End Function
