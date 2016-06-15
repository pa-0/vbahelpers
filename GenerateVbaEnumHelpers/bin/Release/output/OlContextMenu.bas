Attribute VB_Name = "wOlContextMenu"
Function OlContextMenuFromString(value As String) As OlContextMenu
    If IsNumeric(value) Then
        OlContextMenuFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olItemContextMenu": OlContextMenuFromString = olItemContextMenu
        Case "olViewContextMenu": OlContextMenuFromString = olViewContextMenu
        Case "olFolderContextMenu": OlContextMenuFromString = olFolderContextMenu
        Case "olAttachmentContextMenu": OlContextMenuFromString = olAttachmentContextMenu
        Case "olStoreContextMenu": OlContextMenuFromString = olStoreContextMenu
        Case "olShortcutContextMenu": OlContextMenuFromString = olShortcutContextMenu
    End Select
End Function

Function OlContextMenuToString(value As OlContextMenu) As String
    Select Case value
        Case olItemContextMenu: OlContextMenuToString = "olItemContextMenu"
        Case olViewContextMenu: OlContextMenuToString = "olViewContextMenu"
        Case olFolderContextMenu: OlContextMenuToString = "olFolderContextMenu"
        Case olAttachmentContextMenu: OlContextMenuToString = "olAttachmentContextMenu"
        Case olStoreContextMenu: OlContextMenuToString = "olStoreContextMenu"
        Case olShortcutContextMenu: OlContextMenuToString = "olShortcutContextMenu"
    End Select
End Function
