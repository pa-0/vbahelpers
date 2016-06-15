Attribute VB_Name = "wPbVerticalPictureLocking"
Function PbVerticalPictureLockingFromString(value As String) As PbVerticalPictureLocking
    If IsNumeric(value) Then
        PbVerticalPictureLockingFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbVerticalLockingNone": PbVerticalPictureLockingFromString = pbVerticalLockingNone
        Case "pbVerticalLockingTop": PbVerticalPictureLockingFromString = pbVerticalLockingTop
        Case "pbVerticalLockingBottom": PbVerticalPictureLockingFromString = pbVerticalLockingBottom
        Case "pbVerticalLockingStretch": PbVerticalPictureLockingFromString = pbVerticalLockingStretch
    End Select
End Function

Function PbVerticalPictureLockingToString(value As PbVerticalPictureLocking) As String
    Select Case value
        Case pbVerticalLockingNone: PbVerticalPictureLockingToString = "pbVerticalLockingNone"
        Case pbVerticalLockingTop: PbVerticalPictureLockingToString = "pbVerticalLockingTop"
        Case pbVerticalLockingBottom: PbVerticalPictureLockingToString = "pbVerticalLockingBottom"
        Case pbVerticalLockingStretch: PbVerticalPictureLockingToString = "pbVerticalLockingStretch"
    End Select
End Function
