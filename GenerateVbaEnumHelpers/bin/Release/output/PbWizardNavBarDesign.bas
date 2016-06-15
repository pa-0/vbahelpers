Attribute VB_Name = "wPbWizardNavBarDesign"
Function PbWizardNavBarDesignFromString(value As String) As PbWizardNavBarDesign
    If IsNumeric(value) Then
        PbWizardNavBarDesignFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbnbDesignRectangle": PbWizardNavBarDesignFromString = pbnbDesignRectangle
        Case "pbnbDesignAmbient": PbWizardNavBarDesignFromString = pbnbDesignAmbient
        Case "pbnbDesignCapsule": PbWizardNavBarDesignFromString = pbnbDesignCapsule
        Case "pbnbDesignTopDrawer": PbWizardNavBarDesignFromString = pbnbDesignTopDrawer
        Case "pbnbDesignOutline": PbWizardNavBarDesignFromString = pbnbDesignOutline
        Case "pbnbDesignRadius": PbWizardNavBarDesignFromString = pbnbDesignRadius
        Case "pbnbDesignOffset": PbWizardNavBarDesignFromString = pbnbDesignOffset
        Case "pbnbDesignDimension": PbWizardNavBarDesignFromString = pbnbDesignDimension
        Case "pbnbDesignDottedArrow": PbWizardNavBarDesignFromString = pbnbDesignDottedArrow
        Case "pbnbDesignHollowArrow": PbWizardNavBarDesignFromString = pbnbDesignHollowArrow
        Case "pbnbDesignBracket": PbWizardNavBarDesignFromString = pbnbDesignBracket
        Case "pbnbDesignEnclosedArrow": PbWizardNavBarDesignFromString = pbnbDesignEnclosedArrow
        Case "pbnbDesignCounter": PbWizardNavBarDesignFromString = pbnbDesignCounter
        Case "pbnbDesignEndCap": PbWizardNavBarDesignFromString = pbnbDesignEndCap
        Case "pbnbDesignCornice": PbWizardNavBarDesignFromString = pbnbDesignCornice
        Case "pbnbDesignStaff": PbWizardNavBarDesignFromString = pbnbDesignStaff
        Case "pbnbDesignEdge": PbWizardNavBarDesignFromString = pbnbDesignEdge
        Case "pbnbDesignTopLine": PbWizardNavBarDesignFromString = pbnbDesignTopLine
        Case "pbnbDesignUnderscore": PbWizardNavBarDesignFromString = pbnbDesignUnderscore
        Case "pbnbDesignBulletStaff": PbWizardNavBarDesignFromString = pbnbDesignBulletStaff
        Case "pbnbDesignTopBar": PbWizardNavBarDesignFromString = pbnbDesignTopBar
        Case "pbnbDesignKeyPunch": PbWizardNavBarDesignFromString = pbnbDesignKeyPunch
        Case "pbnbDesignRoundBullet": PbWizardNavBarDesignFromString = pbnbDesignRoundBullet
        Case "pbnbDesignSquareBullet": PbWizardNavBarDesignFromString = pbnbDesignSquareBullet
        Case "pbnbDesignWatermark": PbWizardNavBarDesignFromString = pbnbDesignWatermark
        Case "pbnbDesignBaseline": PbWizardNavBarDesignFromString = pbnbDesignBaseline
    End Select
End Function

Function PbWizardNavBarDesignToString(value As PbWizardNavBarDesign) As String
    Select Case value
        Case pbnbDesignRectangle: PbWizardNavBarDesignToString = "pbnbDesignRectangle"
        Case pbnbDesignAmbient: PbWizardNavBarDesignToString = "pbnbDesignAmbient"
        Case pbnbDesignCapsule: PbWizardNavBarDesignToString = "pbnbDesignCapsule"
        Case pbnbDesignTopDrawer: PbWizardNavBarDesignToString = "pbnbDesignTopDrawer"
        Case pbnbDesignOutline: PbWizardNavBarDesignToString = "pbnbDesignOutline"
        Case pbnbDesignRadius: PbWizardNavBarDesignToString = "pbnbDesignRadius"
        Case pbnbDesignOffset: PbWizardNavBarDesignToString = "pbnbDesignOffset"
        Case pbnbDesignDimension: PbWizardNavBarDesignToString = "pbnbDesignDimension"
        Case pbnbDesignDottedArrow: PbWizardNavBarDesignToString = "pbnbDesignDottedArrow"
        Case pbnbDesignHollowArrow: PbWizardNavBarDesignToString = "pbnbDesignHollowArrow"
        Case pbnbDesignBracket: PbWizardNavBarDesignToString = "pbnbDesignBracket"
        Case pbnbDesignEnclosedArrow: PbWizardNavBarDesignToString = "pbnbDesignEnclosedArrow"
        Case pbnbDesignCounter: PbWizardNavBarDesignToString = "pbnbDesignCounter"
        Case pbnbDesignEndCap: PbWizardNavBarDesignToString = "pbnbDesignEndCap"
        Case pbnbDesignCornice: PbWizardNavBarDesignToString = "pbnbDesignCornice"
        Case pbnbDesignStaff: PbWizardNavBarDesignToString = "pbnbDesignStaff"
        Case pbnbDesignEdge: PbWizardNavBarDesignToString = "pbnbDesignEdge"
        Case pbnbDesignTopLine: PbWizardNavBarDesignToString = "pbnbDesignTopLine"
        Case pbnbDesignUnderscore: PbWizardNavBarDesignToString = "pbnbDesignUnderscore"
        Case pbnbDesignBulletStaff: PbWizardNavBarDesignToString = "pbnbDesignBulletStaff"
        Case pbnbDesignTopBar: PbWizardNavBarDesignToString = "pbnbDesignTopBar"
        Case pbnbDesignKeyPunch: PbWizardNavBarDesignToString = "pbnbDesignKeyPunch"
        Case pbnbDesignRoundBullet: PbWizardNavBarDesignToString = "pbnbDesignRoundBullet"
        Case pbnbDesignSquareBullet: PbWizardNavBarDesignToString = "pbnbDesignSquareBullet"
        Case pbnbDesignWatermark: PbWizardNavBarDesignToString = "pbnbDesignWatermark"
        Case pbnbDesignBaseline: PbWizardNavBarDesignToString = "pbnbDesignBaseline"
    End Select
End Function
