Attribute VB_Name = "wOlDownloadState"
Function OlDownloadStateFromString(value As String) As OlDownloadState
    If IsNumeric(value) Then
        OlDownloadStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olHeaderOnly": OlDownloadStateFromString = olHeaderOnly
        Case "olFullItem": OlDownloadStateFromString = olFullItem
    End Select
End Function

Function OlDownloadStateToString(value As OlDownloadState) As String
    Select Case value
        Case olHeaderOnly: OlDownloadStateToString = "olHeaderOnly"
        Case olFullItem: OlDownloadStateToString = "olFullItem"
    End Select
End Function
