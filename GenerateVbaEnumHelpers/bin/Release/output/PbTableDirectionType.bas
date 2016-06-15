Attribute VB_Name = "wPbTableDirectionType"
Function PbTableDirectionTypeFromString(value As String) As PbTableDirectionType
    If IsNumeric(value) Then
        PbTableDirectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTableDirectionLeftToRight": PbTableDirectionTypeFromString = pbTableDirectionLeftToRight
        Case "pbTableDirectionRightToLeft": PbTableDirectionTypeFromString = pbTableDirectionRightToLeft
    End Select
End Function

Function PbTableDirectionTypeToString(value As PbTableDirectionType) As String
    Select Case value
        Case pbTableDirectionLeftToRight: PbTableDirectionTypeToString = "pbTableDirectionLeftToRight"
        Case pbTableDirectionRightToLeft: PbTableDirectionTypeToString = "pbTableDirectionRightToLeft"
    End Select
End Function
