Attribute VB_Name = "wPbTableAutoFormatType"
Function PbTableAutoFormatTypeFromString(value As String) As PbTableAutoFormatType
    If IsNumeric(value) Then
        PbTableAutoFormatTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTableAutoFormatCheckbookRegister": PbTableAutoFormatTypeFromString = pbTableAutoFormatCheckbookRegister
        Case "pbTableAutoFormatList1": PbTableAutoFormatTypeFromString = pbTableAutoFormatList1
        Case "pbTableAutoFormatList2": PbTableAutoFormatTypeFromString = pbTableAutoFormatList2
        Case "pbTableAutoFormatList3": PbTableAutoFormatTypeFromString = pbTableAutoFormatList3
        Case "pbTableAutoFormatList4": PbTableAutoFormatTypeFromString = pbTableAutoFormatList4
        Case "pbTableAutoFormatList5": PbTableAutoFormatTypeFromString = pbTableAutoFormatList5
        Case "pbTableAutoFormatList6": PbTableAutoFormatTypeFromString = pbTableAutoFormatList6
        Case "pbTableAutoFormatList7": PbTableAutoFormatTypeFromString = pbTableAutoFormatList7
        Case "pbTableAutoFormatListWithTitle1": PbTableAutoFormatTypeFromString = pbTableAutoFormatListWithTitle1
        Case "pbTableAutoFormatListWithTitle2": PbTableAutoFormatTypeFromString = pbTableAutoFormatListWithTitle2
        Case "pbTableAutoFormatListWithTitle3": PbTableAutoFormatTypeFromString = pbTableAutoFormatListWithTitle3
        Case "pbTableAutoFormatNumbers1": PbTableAutoFormatTypeFromString = pbTableAutoFormatNumbers1
        Case "pbTableAutoFormatNumbers2": PbTableAutoFormatTypeFromString = pbTableAutoFormatNumbers2
        Case "pbTableAutoFormatNumbers3": PbTableAutoFormatTypeFromString = pbTableAutoFormatNumbers3
        Case "pbTableAutoFormatNumbers4": PbTableAutoFormatTypeFromString = pbTableAutoFormatNumbers4
        Case "pbTableAutoFormatNumbers5": PbTableAutoFormatTypeFromString = pbTableAutoFormatNumbers5
        Case "pbTableAutoFormatNumbers6": PbTableAutoFormatTypeFromString = pbTableAutoFormatNumbers6
        Case "pbTableAutoFormatTableOfContents1": PbTableAutoFormatTypeFromString = pbTableAutoFormatTableOfContents1
        Case "pbTableAutoFormatTableOfContents2": PbTableAutoFormatTypeFromString = pbTableAutoFormatTableOfContents2
        Case "pbTableAutoFormatTableOfContents3": PbTableAutoFormatTypeFromString = pbTableAutoFormatTableOfContents3
        Case "pbTableAutoFormatCheckerboard": PbTableAutoFormatTypeFromString = pbTableAutoFormatCheckerboard
        Case "pbTableAutoFormatDefault": PbTableAutoFormatTypeFromString = pbTableAutoFormatDefault
        Case "pbTableAutoFormatNone": PbTableAutoFormatTypeFromString = pbTableAutoFormatNone
        Case "pbTableAutoFormatMixed": PbTableAutoFormatTypeFromString = pbTableAutoFormatMixed
    End Select
End Function

Function PbTableAutoFormatTypeToString(value As PbTableAutoFormatType) As String
    Select Case value
        Case pbTableAutoFormatCheckbookRegister: PbTableAutoFormatTypeToString = "pbTableAutoFormatCheckbookRegister"
        Case pbTableAutoFormatList1: PbTableAutoFormatTypeToString = "pbTableAutoFormatList1"
        Case pbTableAutoFormatList2: PbTableAutoFormatTypeToString = "pbTableAutoFormatList2"
        Case pbTableAutoFormatList3: PbTableAutoFormatTypeToString = "pbTableAutoFormatList3"
        Case pbTableAutoFormatList4: PbTableAutoFormatTypeToString = "pbTableAutoFormatList4"
        Case pbTableAutoFormatList5: PbTableAutoFormatTypeToString = "pbTableAutoFormatList5"
        Case pbTableAutoFormatList6: PbTableAutoFormatTypeToString = "pbTableAutoFormatList6"
        Case pbTableAutoFormatList7: PbTableAutoFormatTypeToString = "pbTableAutoFormatList7"
        Case pbTableAutoFormatListWithTitle1: PbTableAutoFormatTypeToString = "pbTableAutoFormatListWithTitle1"
        Case pbTableAutoFormatListWithTitle2: PbTableAutoFormatTypeToString = "pbTableAutoFormatListWithTitle2"
        Case pbTableAutoFormatListWithTitle3: PbTableAutoFormatTypeToString = "pbTableAutoFormatListWithTitle3"
        Case pbTableAutoFormatNumbers1: PbTableAutoFormatTypeToString = "pbTableAutoFormatNumbers1"
        Case pbTableAutoFormatNumbers2: PbTableAutoFormatTypeToString = "pbTableAutoFormatNumbers2"
        Case pbTableAutoFormatNumbers3: PbTableAutoFormatTypeToString = "pbTableAutoFormatNumbers3"
        Case pbTableAutoFormatNumbers4: PbTableAutoFormatTypeToString = "pbTableAutoFormatNumbers4"
        Case pbTableAutoFormatNumbers5: PbTableAutoFormatTypeToString = "pbTableAutoFormatNumbers5"
        Case pbTableAutoFormatNumbers6: PbTableAutoFormatTypeToString = "pbTableAutoFormatNumbers6"
        Case pbTableAutoFormatTableOfContents1: PbTableAutoFormatTypeToString = "pbTableAutoFormatTableOfContents1"
        Case pbTableAutoFormatTableOfContents2: PbTableAutoFormatTypeToString = "pbTableAutoFormatTableOfContents2"
        Case pbTableAutoFormatTableOfContents3: PbTableAutoFormatTypeToString = "pbTableAutoFormatTableOfContents3"
        Case pbTableAutoFormatCheckerboard: PbTableAutoFormatTypeToString = "pbTableAutoFormatCheckerboard"
        Case pbTableAutoFormatDefault: PbTableAutoFormatTypeToString = "pbTableAutoFormatDefault"
        Case pbTableAutoFormatNone: PbTableAutoFormatTypeToString = "pbTableAutoFormatNone"
        Case pbTableAutoFormatMixed: PbTableAutoFormatTypeToString = "pbTableAutoFormatMixed"
    End Select
End Function
