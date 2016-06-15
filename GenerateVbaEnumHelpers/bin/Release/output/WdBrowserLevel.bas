Attribute VB_Name = "wWdBrowserLevel"
Function WdBrowserLevelFromString(value As String) As WdBrowserLevel
    If IsNumeric(value) Then
        WdBrowserLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBrowserLevelV4": WdBrowserLevelFromString = wdBrowserLevelV4
        Case "wdBrowserLevelMicrosoftInternetExplorer5": WdBrowserLevelFromString = wdBrowserLevelMicrosoftInternetExplorer5
        Case "wdBrowserLevelMicrosoftInternetExplorer6": WdBrowserLevelFromString = wdBrowserLevelMicrosoftInternetExplorer6
    End Select
End Function

Function WdBrowserLevelToString(value As WdBrowserLevel) As String
    Select Case value
        Case wdBrowserLevelV4: WdBrowserLevelToString = "wdBrowserLevelV4"
        Case wdBrowserLevelMicrosoftInternetExplorer5: WdBrowserLevelToString = "wdBrowserLevelMicrosoftInternetExplorer5"
        Case wdBrowserLevelMicrosoftInternetExplorer6: WdBrowserLevelToString = "wdBrowserLevelMicrosoftInternetExplorer6"
    End Select
End Function
