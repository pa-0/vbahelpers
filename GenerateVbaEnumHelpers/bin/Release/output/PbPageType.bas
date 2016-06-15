Attribute VB_Name = "wPbPageType"
Function PbPageTypeFromString(value As String) As PbPageType
    If IsNumeric(value) Then
        PbPageTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPageLeftPage": PbPageTypeFromString = pbPageLeftPage
        Case "pbPageRightPage": PbPageTypeFromString = pbPageRightPage
        Case "pbPageScratchPage": PbPageTypeFromString = pbPageScratchPage
        Case "pbPageMasterPage": PbPageTypeFromString = pbPageMasterPage
    End Select
End Function

Function PbPageTypeToString(value As PbPageType) As String
    Select Case value
        Case pbPageLeftPage: PbPageTypeToString = "pbPageLeftPage"
        Case pbPageRightPage: PbPageTypeToString = "pbPageRightPage"
        Case pbPageScratchPage: PbPageTypeToString = "pbPageScratchPage"
        Case pbPageMasterPage: PbPageTypeToString = "pbPageMasterPage"
    End Select
End Function
