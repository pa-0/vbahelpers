Attribute VB_Name = "wMsoButtonStyle"
Function MsoButtonStyleFromString(value As String) As MsoButtonStyle
    If IsNumeric(value) Then
        MsoButtonStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoButtonAutomatic": MsoButtonStyleFromString = msoButtonAutomatic
        Case "msoButtonIcon": MsoButtonStyleFromString = msoButtonIcon
        Case "msoButtonCaption": MsoButtonStyleFromString = msoButtonCaption
        Case "msoButtonIconAndCaption": MsoButtonStyleFromString = msoButtonIconAndCaption
        Case "msoButtonIconAndWrapCaption": MsoButtonStyleFromString = msoButtonIconAndWrapCaption
        Case "msoButtonIconAndCaptionBelow": MsoButtonStyleFromString = msoButtonIconAndCaptionBelow
        Case "msoButtonWrapCaption": MsoButtonStyleFromString = msoButtonWrapCaption
        Case "msoButtonIconAndWrapCaptionBelow": MsoButtonStyleFromString = msoButtonIconAndWrapCaptionBelow
    End Select
End Function

Function MsoButtonStyleToString(value As MsoButtonStyle) As String
    Select Case value
        Case msoButtonAutomatic: MsoButtonStyleToString = "msoButtonAutomatic"
        Case msoButtonIcon: MsoButtonStyleToString = "msoButtonIcon"
        Case msoButtonCaption: MsoButtonStyleToString = "msoButtonCaption"
        Case msoButtonIconAndCaption: MsoButtonStyleToString = "msoButtonIconAndCaption"
        Case msoButtonIconAndWrapCaption: MsoButtonStyleToString = "msoButtonIconAndWrapCaption"
        Case msoButtonIconAndCaptionBelow: MsoButtonStyleToString = "msoButtonIconAndCaptionBelow"
        Case msoButtonWrapCaption: MsoButtonStyleToString = "msoButtonWrapCaption"
        Case msoButtonIconAndWrapCaptionBelow: MsoButtonStyleToString = "msoButtonIconAndWrapCaptionBelow"
    End Select
End Function
