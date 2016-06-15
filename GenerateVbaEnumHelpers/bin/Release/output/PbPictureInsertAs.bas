Attribute VB_Name = "wPbPictureInsertAs"
Function PbPictureInsertAsFromString(value As String) As PbPictureInsertAs
    If IsNumeric(value) Then
        PbPictureInsertAsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPictureInsertAsEmbedded": PbPictureInsertAsFromString = pbPictureInsertAsEmbedded
        Case "pbPictureInsertAsLinked": PbPictureInsertAsFromString = pbPictureInsertAsLinked
        Case "pbPictureInsertAsOriginalState": PbPictureInsertAsFromString = pbPictureInsertAsOriginalState
    End Select
End Function

Function PbPictureInsertAsToString(value As PbPictureInsertAs) As String
    Select Case value
        Case pbPictureInsertAsEmbedded: PbPictureInsertAsToString = "pbPictureInsertAsEmbedded"
        Case pbPictureInsertAsLinked: PbPictureInsertAsToString = "pbPictureInsertAsLinked"
        Case pbPictureInsertAsOriginalState: PbPictureInsertAsToString = "pbPictureInsertAsOriginalState"
    End Select
End Function
