Attribute VB_Name = "wMsoBulletType"
Function MsoBulletTypeFromString(value As String) As MsoBulletType
    If IsNumeric(value) Then
        MsoBulletTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBulletNone": MsoBulletTypeFromString = msoBulletNone
        Case "msoBulletUnnumbered": MsoBulletTypeFromString = msoBulletUnnumbered
        Case "msoBulletNumbered": MsoBulletTypeFromString = msoBulletNumbered
        Case "msoBulletPicture": MsoBulletTypeFromString = msoBulletPicture
        Case "msoBulletMixed": MsoBulletTypeFromString = msoBulletMixed
    End Select
End Function

Function MsoBulletTypeToString(value As MsoBulletType) As String
    Select Case value
        Case msoBulletNone: MsoBulletTypeToString = "msoBulletNone"
        Case msoBulletUnnumbered: MsoBulletTypeToString = "msoBulletUnnumbered"
        Case msoBulletNumbered: MsoBulletTypeToString = "msoBulletNumbered"
        Case msoBulletPicture: MsoBulletTypeToString = "msoBulletPicture"
        Case msoBulletMixed: MsoBulletTypeToString = "msoBulletMixed"
    End Select
End Function
