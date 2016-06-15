Attribute VB_Name = "wWdDefaultFilePath"
Function WdDefaultFilePathFromString(value As String) As WdDefaultFilePath
    If IsNumeric(value) Then
        WdDefaultFilePathFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDocumentsPath": WdDefaultFilePathFromString = wdDocumentsPath
        Case "wdPicturesPath": WdDefaultFilePathFromString = wdPicturesPath
        Case "wdUserTemplatesPath": WdDefaultFilePathFromString = wdUserTemplatesPath
        Case "wdWorkgroupTemplatesPath": WdDefaultFilePathFromString = wdWorkgroupTemplatesPath
        Case "wdUserOptionsPath": WdDefaultFilePathFromString = wdUserOptionsPath
        Case "wdAutoRecoverPath": WdDefaultFilePathFromString = wdAutoRecoverPath
        Case "wdToolsPath": WdDefaultFilePathFromString = wdToolsPath
        Case "wdTutorialPath": WdDefaultFilePathFromString = wdTutorialPath
        Case "wdStartupPath": WdDefaultFilePathFromString = wdStartupPath
        Case "wdProgramPath": WdDefaultFilePathFromString = wdProgramPath
        Case "wdGraphicsFiltersPath": WdDefaultFilePathFromString = wdGraphicsFiltersPath
        Case "wdTextConvertersPath": WdDefaultFilePathFromString = wdTextConvertersPath
        Case "wdProofingToolsPath": WdDefaultFilePathFromString = wdProofingToolsPath
        Case "wdTempFilePath": WdDefaultFilePathFromString = wdTempFilePath
        Case "wdCurrentFolderPath": WdDefaultFilePathFromString = wdCurrentFolderPath
        Case "wdStyleGalleryPath": WdDefaultFilePathFromString = wdStyleGalleryPath
        Case "wdBorderArtPath": WdDefaultFilePathFromString = wdBorderArtPath
    End Select
End Function

Function WdDefaultFilePathToString(value As WdDefaultFilePath) As String
    Select Case value
        Case wdDocumentsPath: WdDefaultFilePathToString = "wdDocumentsPath"
        Case wdPicturesPath: WdDefaultFilePathToString = "wdPicturesPath"
        Case wdUserTemplatesPath: WdDefaultFilePathToString = "wdUserTemplatesPath"
        Case wdWorkgroupTemplatesPath: WdDefaultFilePathToString = "wdWorkgroupTemplatesPath"
        Case wdUserOptionsPath: WdDefaultFilePathToString = "wdUserOptionsPath"
        Case wdAutoRecoverPath: WdDefaultFilePathToString = "wdAutoRecoverPath"
        Case wdToolsPath: WdDefaultFilePathToString = "wdToolsPath"
        Case wdTutorialPath: WdDefaultFilePathToString = "wdTutorialPath"
        Case wdStartupPath: WdDefaultFilePathToString = "wdStartupPath"
        Case wdProgramPath: WdDefaultFilePathToString = "wdProgramPath"
        Case wdGraphicsFiltersPath: WdDefaultFilePathToString = "wdGraphicsFiltersPath"
        Case wdTextConvertersPath: WdDefaultFilePathToString = "wdTextConvertersPath"
        Case wdProofingToolsPath: WdDefaultFilePathToString = "wdProofingToolsPath"
        Case wdTempFilePath: WdDefaultFilePathToString = "wdTempFilePath"
        Case wdCurrentFolderPath: WdDefaultFilePathToString = "wdCurrentFolderPath"
        Case wdStyleGalleryPath: WdDefaultFilePathToString = "wdStyleGalleryPath"
        Case wdBorderArtPath: WdDefaultFilePathToString = "wdBorderArtPath"
    End Select
End Function
