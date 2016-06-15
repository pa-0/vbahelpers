Attribute VB_Name = "wOlPane"
Function OlPaneFromString(value As String) As OlPane
    If IsNumeric(value) Then
        OlPaneFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olOutlookBar": OlPaneFromString = olOutlookBar
        Case "olFolderList": OlPaneFromString = olFolderList
        Case "olPreview": OlPaneFromString = olPreview
        Case "olNavigationPane": OlPaneFromString = olNavigationPane
        Case "olToDoBar": OlPaneFromString = olToDoBar
    End Select
End Function

Function OlPaneToString(value As OlPane) As String
    Select Case value
        Case olOutlookBar: OlPaneToString = "olOutlookBar"
        Case olFolderList: OlPaneToString = "olFolderList"
        Case olPreview: OlPaneToString = "olPreview"
        Case olNavigationPane: OlPaneToString = "olNavigationPane"
        Case olToDoBar: OlPaneToString = "olToDoBar"
    End Select
End Function
