Attribute VB_Name = "wMsoTargetBrowser"
Function MsoTargetBrowserFromString(value As String) As MsoTargetBrowser
    If IsNumeric(value) Then
        MsoTargetBrowserFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTargetBrowserV3": MsoTargetBrowserFromString = msoTargetBrowserV3
        Case "msoTargetBrowserV4": MsoTargetBrowserFromString = msoTargetBrowserV4
        Case "msoTargetBrowserIE4": MsoTargetBrowserFromString = msoTargetBrowserIE4
        Case "msoTargetBrowserIE5": MsoTargetBrowserFromString = msoTargetBrowserIE5
        Case "msoTargetBrowserIE6": MsoTargetBrowserFromString = msoTargetBrowserIE6
    End Select
End Function

Function MsoTargetBrowserToString(value As MsoTargetBrowser) As String
    Select Case value
        Case msoTargetBrowserV3: MsoTargetBrowserToString = "msoTargetBrowserV3"
        Case msoTargetBrowserV4: MsoTargetBrowserToString = "msoTargetBrowserV4"
        Case msoTargetBrowserIE4: MsoTargetBrowserToString = "msoTargetBrowserIE4"
        Case msoTargetBrowserIE5: MsoTargetBrowserToString = "msoTargetBrowserIE5"
        Case msoTargetBrowserIE6: MsoTargetBrowserToString = "msoTargetBrowserIE6"
    End Select
End Function
