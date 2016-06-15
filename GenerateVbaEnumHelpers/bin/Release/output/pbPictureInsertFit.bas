Attribute VB_Name = "wpbPictureInsertFit"
Function pbPictureInsertFitFromString(value As String) As pbPictureInsertFit
    If IsNumeric(value) Then
        pbPictureInsertFitFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFit": pbPictureInsertFitFromString = pbFit
        Case "pbFill": pbPictureInsertFitFromString = pbFill
    End Select
End Function

Function pbPictureInsertFitToString(value As pbPictureInsertFit) As String
    Select Case value
        Case pbFit: pbPictureInsertFitToString = "pbFit"
        Case pbFill: pbPictureInsertFitToString = "pbFill"
    End Select
End Function
