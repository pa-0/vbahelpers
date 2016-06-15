Attribute VB_Name = "wPbWizardNavBarButtonStyle"
Function PbWizardNavBarButtonStyleFromString(value As String) As PbWizardNavBarButtonStyle
    If IsNumeric(value) Then
        PbWizardNavBarButtonStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbnbButtonStyleSmall": PbWizardNavBarButtonStyleFromString = pbnbButtonStyleSmall
        Case "pbnbButtonStyleLarge": PbWizardNavBarButtonStyleFromString = pbnbButtonStyleLarge
        Case "pbnbButtonStyleText": PbWizardNavBarButtonStyleFromString = pbnbButtonStyleText
    End Select
End Function

Function PbWizardNavBarButtonStyleToString(value As PbWizardNavBarButtonStyle) As String
    Select Case value
        Case pbnbButtonStyleSmall: PbWizardNavBarButtonStyleToString = "pbnbButtonStyleSmall"
        Case pbnbButtonStyleLarge: PbWizardNavBarButtonStyleToString = "pbnbButtonStyleLarge"
        Case pbnbButtonStyleText: PbWizardNavBarButtonStyleToString = "pbnbButtonStyleText"
    End Select
End Function
