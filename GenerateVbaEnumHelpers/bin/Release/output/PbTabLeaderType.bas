Attribute VB_Name = "wPbTabLeaderType"
Function PbTabLeaderTypeFromString(value As String) As PbTabLeaderType
    If IsNumeric(value) Then
        PbTabLeaderTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTabLeaderNone": PbTabLeaderTypeFromString = pbTabLeaderNone
        Case "pbTabLeaderDot": PbTabLeaderTypeFromString = pbTabLeaderDot
        Case "pbTabLeaderDashes": PbTabLeaderTypeFromString = pbTabLeaderDashes
        Case "pbTabLeaderLine": PbTabLeaderTypeFromString = pbTabLeaderLine
        Case "pbTabLeaderBullet": PbTabLeaderTypeFromString = pbTabLeaderBullet
    End Select
End Function

Function PbTabLeaderTypeToString(value As PbTabLeaderType) As String
    Select Case value
        Case pbTabLeaderNone: PbTabLeaderTypeToString = "pbTabLeaderNone"
        Case pbTabLeaderDot: PbTabLeaderTypeToString = "pbTabLeaderDot"
        Case pbTabLeaderDashes: PbTabLeaderTypeToString = "pbTabLeaderDashes"
        Case pbTabLeaderLine: PbTabLeaderTypeToString = "pbTabLeaderLine"
        Case pbTabLeaderBullet: PbTabLeaderTypeToString = "pbTabLeaderBullet"
    End Select
End Function
