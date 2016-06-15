Attribute VB_Name = "wPbDirectionType"
Function PbDirectionTypeFromString(value As String) As PbDirectionType
    If IsNumeric(value) Then
        PbDirectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbDirectionLeftToRight": PbDirectionTypeFromString = pbDirectionLeftToRight
        Case "pbDirectionRightToLeft": PbDirectionTypeFromString = pbDirectionRightToLeft
    End Select
End Function

Function PbDirectionTypeToString(value As PbDirectionType) As String
    Select Case value
        Case pbDirectionLeftToRight: PbDirectionTypeToString = "pbDirectionLeftToRight"
        Case pbDirectionRightToLeft: PbDirectionTypeToString = "pbDirectionRightToLeft"
    End Select
End Function
