Attribute VB_Name = "wPbParagraphAlignmentType"
Function PbParagraphAlignmentTypeFromString(value As String) As PbParagraphAlignmentType
    If IsNumeric(value) Then
        PbParagraphAlignmentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbParagraphAlignmentLeft": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentLeft
        Case "pbParagraphAlignmentCenter": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentCenter
        Case "pbParagraphAlignmentRight": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentRight
        Case "pbParagraphAlignmentInterWord": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentInterWord
        Case "pbParagraphAlignmentDistribute": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentDistribute
        Case "pbParagraphAlignmentDistributeEastAsia": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentDistributeEastAsia
        Case "pbParagraphAlignmentJustified": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentJustified
        Case "pbParagraphAlignmentInterIdeograph": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentInterIdeograph
        Case "pbParagraphAlignmentInterCluster": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentInterCluster
        Case "pbParagraphAlignmentDistributeAll": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentDistributeAll
        Case "pbParagraphAlignmentDistributeCenterLast": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentDistributeCenterLast
        Case "pbParagraphAlignmentKashida": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentKashida
        Case "pbParagraphAlignmentMixed": PbParagraphAlignmentTypeFromString = pbParagraphAlignmentMixed
    End Select
End Function

Function PbParagraphAlignmentTypeToString(value As PbParagraphAlignmentType) As String
    Select Case value
        Case pbParagraphAlignmentLeft: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentLeft"
        Case pbParagraphAlignmentCenter: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentCenter"
        Case pbParagraphAlignmentRight: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentRight"
        Case pbParagraphAlignmentInterWord: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentInterWord"
        Case pbParagraphAlignmentDistribute: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentDistribute"
        Case pbParagraphAlignmentDistributeEastAsia: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentDistributeEastAsia"
        Case pbParagraphAlignmentJustified: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentJustified"
        Case pbParagraphAlignmentInterIdeograph: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentInterIdeograph"
        Case pbParagraphAlignmentInterCluster: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentInterCluster"
        Case pbParagraphAlignmentDistributeAll: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentDistributeAll"
        Case pbParagraphAlignmentDistributeCenterLast: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentDistributeCenterLast"
        Case pbParagraphAlignmentKashida: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentKashida"
        Case pbParagraphAlignmentMixed: PbParagraphAlignmentTypeToString = "pbParagraphAlignmentMixed"
    End Select
End Function
