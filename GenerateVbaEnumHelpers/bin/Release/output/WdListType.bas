Attribute VB_Name = "wWdListType"
Function WdListTypeFromString(value As String) As WdListType
    If IsNumeric(value) Then
        WdListTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdListNoNumbering": WdListTypeFromString = wdListNoNumbering
        Case "wdListListNumOnly": WdListTypeFromString = wdListListNumOnly
        Case "wdListBullet": WdListTypeFromString = wdListBullet
        Case "wdListSimpleNumbering": WdListTypeFromString = wdListSimpleNumbering
        Case "wdListOutlineNumbering": WdListTypeFromString = wdListOutlineNumbering
        Case "wdListMixedNumbering": WdListTypeFromString = wdListMixedNumbering
        Case "wdListPictureBullet": WdListTypeFromString = wdListPictureBullet
    End Select
End Function

Function WdListTypeToString(value As WdListType) As String
    Select Case value
        Case wdListNoNumbering: WdListTypeToString = "wdListNoNumbering"
        Case wdListListNumOnly: WdListTypeToString = "wdListListNumOnly"
        Case wdListBullet: WdListTypeToString = "wdListBullet"
        Case wdListSimpleNumbering: WdListTypeToString = "wdListSimpleNumbering"
        Case wdListOutlineNumbering: WdListTypeToString = "wdListOutlineNumbering"
        Case wdListMixedNumbering: WdListTypeToString = "wdListMixedNumbering"
        Case wdListPictureBullet: WdListTypeToString = "wdListPictureBullet"
    End Select
End Function
