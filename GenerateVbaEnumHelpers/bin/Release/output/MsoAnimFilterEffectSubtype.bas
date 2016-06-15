Attribute VB_Name = "wMsoAnimFilterEffectSubtype"
Function MsoAnimFilterEffectSubtypeFromString(value As String) As MsoAnimFilterEffectSubtype
    If IsNumeric(value) Then
        MsoAnimFilterEffectSubtypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimFilterEffectSubtypeNone": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeNone
        Case "msoAnimFilterEffectSubtypeInVertical": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeInVertical
        Case "msoAnimFilterEffectSubtypeOutVertical": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeOutVertical
        Case "msoAnimFilterEffectSubtypeInHorizontal": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeInHorizontal
        Case "msoAnimFilterEffectSubtypeOutHorizontal": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeOutHorizontal
        Case "msoAnimFilterEffectSubtypeHorizontal": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeHorizontal
        Case "msoAnimFilterEffectSubtypeVertical": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeVertical
        Case "msoAnimFilterEffectSubtypeIn": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeIn
        Case "msoAnimFilterEffectSubtypeOut": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeOut
        Case "msoAnimFilterEffectSubtypeAcross": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeAcross
        Case "msoAnimFilterEffectSubtypeFromLeft": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeFromLeft
        Case "msoAnimFilterEffectSubtypeFromRight": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeFromRight
        Case "msoAnimFilterEffectSubtypeFromTop": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeFromTop
        Case "msoAnimFilterEffectSubtypeFromBottom": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeFromBottom
        Case "msoAnimFilterEffectSubtypeDownLeft": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeDownLeft
        Case "msoAnimFilterEffectSubtypeUpLeft": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeUpLeft
        Case "msoAnimFilterEffectSubtypeDownRight": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeDownRight
        Case "msoAnimFilterEffectSubtypeUpRight": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeUpRight
        Case "msoAnimFilterEffectSubtypeSpokes1": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeSpokes1
        Case "msoAnimFilterEffectSubtypeSpokes2": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeSpokes2
        Case "msoAnimFilterEffectSubtypeSpokes3": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeSpokes3
        Case "msoAnimFilterEffectSubtypeSpokes4": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeSpokes4
        Case "msoAnimFilterEffectSubtypeSpokes8": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeSpokes8
        Case "msoAnimFilterEffectSubtypeLeft": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeLeft
        Case "msoAnimFilterEffectSubtypeRight": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeRight
        Case "msoAnimFilterEffectSubtypeDown": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeDown
        Case "msoAnimFilterEffectSubtypeUp": MsoAnimFilterEffectSubtypeFromString = msoAnimFilterEffectSubtypeUp
    End Select
End Function

Function MsoAnimFilterEffectSubtypeToString(value As MsoAnimFilterEffectSubtype) As String
    Select Case value
        Case msoAnimFilterEffectSubtypeNone: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeNone"
        Case msoAnimFilterEffectSubtypeInVertical: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeInVertical"
        Case msoAnimFilterEffectSubtypeOutVertical: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeOutVertical"
        Case msoAnimFilterEffectSubtypeInHorizontal: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeInHorizontal"
        Case msoAnimFilterEffectSubtypeOutHorizontal: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeOutHorizontal"
        Case msoAnimFilterEffectSubtypeHorizontal: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeHorizontal"
        Case msoAnimFilterEffectSubtypeVertical: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeVertical"
        Case msoAnimFilterEffectSubtypeIn: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeIn"
        Case msoAnimFilterEffectSubtypeOut: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeOut"
        Case msoAnimFilterEffectSubtypeAcross: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeAcross"
        Case msoAnimFilterEffectSubtypeFromLeft: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeFromLeft"
        Case msoAnimFilterEffectSubtypeFromRight: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeFromRight"
        Case msoAnimFilterEffectSubtypeFromTop: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeFromTop"
        Case msoAnimFilterEffectSubtypeFromBottom: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeFromBottom"
        Case msoAnimFilterEffectSubtypeDownLeft: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeDownLeft"
        Case msoAnimFilterEffectSubtypeUpLeft: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeUpLeft"
        Case msoAnimFilterEffectSubtypeDownRight: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeDownRight"
        Case msoAnimFilterEffectSubtypeUpRight: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeUpRight"
        Case msoAnimFilterEffectSubtypeSpokes1: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeSpokes1"
        Case msoAnimFilterEffectSubtypeSpokes2: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeSpokes2"
        Case msoAnimFilterEffectSubtypeSpokes3: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeSpokes3"
        Case msoAnimFilterEffectSubtypeSpokes4: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeSpokes4"
        Case msoAnimFilterEffectSubtypeSpokes8: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeSpokes8"
        Case msoAnimFilterEffectSubtypeLeft: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeLeft"
        Case msoAnimFilterEffectSubtypeRight: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeRight"
        Case msoAnimFilterEffectSubtypeDown: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeDown"
        Case msoAnimFilterEffectSubtypeUp: MsoAnimFilterEffectSubtypeToString = "msoAnimFilterEffectSubtypeUp"
    End Select
End Function
