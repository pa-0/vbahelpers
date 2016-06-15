Attribute VB_Name = "wPbHorizontalPictureLocking"
Function PbHorizontalPictureLockingFromString(value As String) As PbHorizontalPictureLocking
    If IsNumeric(value) Then
        PbHorizontalPictureLockingFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbHorizontalLockingNone": PbHorizontalPictureLockingFromString = pbHorizontalLockingNone
        Case "pbHorizontalLockingLeft": PbHorizontalPictureLockingFromString = pbHorizontalLockingLeft
        Case "pbHorizontalLockingRight": PbHorizontalPictureLockingFromString = pbHorizontalLockingRight
        Case "pbHorizontalLockingStretch": PbHorizontalPictureLockingFromString = pbHorizontalLockingStretch
    End Select
End Function

Function PbHorizontalPictureLockingToString(value As PbHorizontalPictureLocking) As String
    Select Case value
        Case pbHorizontalLockingNone: PbHorizontalPictureLockingToString = "pbHorizontalLockingNone"
        Case pbHorizontalLockingLeft: PbHorizontalPictureLockingToString = "pbHorizontalLockingLeft"
        Case pbHorizontalLockingRight: PbHorizontalPictureLockingToString = "pbHorizontalLockingRight"
        Case pbHorizontalLockingStretch: PbHorizontalPictureLockingToString = "pbHorizontalLockingStretch"
    End Select
End Function
