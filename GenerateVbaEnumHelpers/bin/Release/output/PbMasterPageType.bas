Attribute VB_Name = "wPbMasterPageType"
Function PbMasterPageTypeFromString(value As String) As PbMasterPageType
    If IsNumeric(value) Then
        PbMasterPageTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbMasterPageLeftPage": PbMasterPageTypeFromString = pbMasterPageLeftPage
        Case "pbMasterPageRightPage": PbMasterPageTypeFromString = pbMasterPageRightPage
    End Select
End Function

Function PbMasterPageTypeToString(value As PbMasterPageType) As String
    Select Case value
        Case pbMasterPageLeftPage: PbMasterPageTypeToString = "pbMasterPageLeftPage"
        Case pbMasterPageRightPage: PbMasterPageTypeToString = "pbMasterPageRightPage"
    End Select
End Function
