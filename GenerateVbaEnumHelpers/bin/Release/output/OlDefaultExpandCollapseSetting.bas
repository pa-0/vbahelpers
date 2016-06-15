Attribute VB_Name = "wOlDefaultExpandCollapseSetting"
Function OlDefaultExpandCollapseSettingFromString(value As String) As OlDefaultExpandCollapseSetting
    If IsNumeric(value) Then
        OlDefaultExpandCollapseSettingFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAllExpanded": OlDefaultExpandCollapseSettingFromString = olAllExpanded
        Case "olAllCollapsed": OlDefaultExpandCollapseSettingFromString = olAllCollapsed
        Case "olLastViewed": OlDefaultExpandCollapseSettingFromString = olLastViewed
    End Select
End Function

Function OlDefaultExpandCollapseSettingToString(value As OlDefaultExpandCollapseSetting) As String
    Select Case value
        Case olAllExpanded: OlDefaultExpandCollapseSettingToString = "olAllExpanded"
        Case olAllCollapsed: OlDefaultExpandCollapseSettingToString = "olAllCollapsed"
        Case olLastViewed: OlDefaultExpandCollapseSettingToString = "olLastViewed"
    End Select
End Function
