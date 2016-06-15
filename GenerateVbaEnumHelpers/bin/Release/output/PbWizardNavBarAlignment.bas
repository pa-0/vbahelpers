Attribute VB_Name = "wPbWizardNavBarAlignment"
Function PbWizardNavBarAlignmentFromString(value As String) As PbWizardNavBarAlignment
    If IsNumeric(value) Then
        PbWizardNavBarAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbnbAlignLeft": PbWizardNavBarAlignmentFromString = pbnbAlignLeft
        Case "pbnbAlignCenter": PbWizardNavBarAlignmentFromString = pbnbAlignCenter
        Case "pbnbAlignRight": PbWizardNavBarAlignmentFromString = pbnbAlignRight
    End Select
End Function

Function PbWizardNavBarAlignmentToString(value As PbWizardNavBarAlignment) As String
    Select Case value
        Case pbnbAlignLeft: PbWizardNavBarAlignmentToString = "pbnbAlignLeft"
        Case pbnbAlignCenter: PbWizardNavBarAlignmentToString = "pbnbAlignCenter"
        Case pbnbAlignRight: PbWizardNavBarAlignmentToString = "pbnbAlignRight"
    End Select
End Function
