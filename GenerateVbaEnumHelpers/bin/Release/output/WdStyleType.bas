Attribute VB_Name = "wWdStyleType"
Function WdStyleTypeFromString(value As String) As WdStyleType
    If IsNumeric(value) Then
        WdStyleTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdStyleTypeParagraph": WdStyleTypeFromString = wdStyleTypeParagraph
        Case "wdStyleTypeCharacter": WdStyleTypeFromString = wdStyleTypeCharacter
        Case "wdStyleTypeTable": WdStyleTypeFromString = wdStyleTypeTable
        Case "wdStyleTypeList": WdStyleTypeFromString = wdStyleTypeList
        Case "wdStyleTypeParagraphOnly": WdStyleTypeFromString = wdStyleTypeParagraphOnly
        Case "wdStyleTypeLinked": WdStyleTypeFromString = wdStyleTypeLinked
    End Select
End Function

Function WdStyleTypeToString(value As WdStyleType) As String
    Select Case value
        Case wdStyleTypeParagraph: WdStyleTypeToString = "wdStyleTypeParagraph"
        Case wdStyleTypeCharacter: WdStyleTypeToString = "wdStyleTypeCharacter"
        Case wdStyleTypeTable: WdStyleTypeToString = "wdStyleTypeTable"
        Case wdStyleTypeList: WdStyleTypeToString = "wdStyleTypeList"
        Case wdStyleTypeParagraphOnly: WdStyleTypeToString = "wdStyleTypeParagraphOnly"
        Case wdStyleTypeLinked: WdStyleTypeToString = "wdStyleTypeLinked"
    End Select
End Function
