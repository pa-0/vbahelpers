Attribute VB_Name = "wOlTableContents"
Function OlTableContentsFromString(value As String) As OlTableContents
    If IsNumeric(value) Then
        OlTableContentsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olUserItems": OlTableContentsFromString = olUserItems
        Case "olHiddenItems": OlTableContentsFromString = olHiddenItems
    End Select
End Function

Function OlTableContentsToString(value As OlTableContents) As String
    Select Case value
        Case olUserItems: OlTableContentsToString = "olUserItems"
        Case olHiddenItems: OlTableContentsToString = "olHiddenItems"
    End Select
End Function
