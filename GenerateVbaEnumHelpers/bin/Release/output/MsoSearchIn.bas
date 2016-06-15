Attribute VB_Name = "wMsoSearchIn"
Function MsoSearchInFromString(value As String) As MsoSearchIn
    If IsNumeric(value) Then
        MsoSearchInFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSearchInMyComputer": MsoSearchInFromString = msoSearchInMyComputer
        Case "msoSearchInOutlook": MsoSearchInFromString = msoSearchInOutlook
        Case "msoSearchInMyNetworkPlaces": MsoSearchInFromString = msoSearchInMyNetworkPlaces
        Case "msoSearchInCustom": MsoSearchInFromString = msoSearchInCustom
    End Select
End Function

Function MsoSearchInToString(value As MsoSearchIn) As String
    Select Case value
        Case msoSearchInMyComputer: MsoSearchInToString = "msoSearchInMyComputer"
        Case msoSearchInOutlook: MsoSearchInToString = "msoSearchInOutlook"
        Case msoSearchInMyNetworkPlaces: MsoSearchInToString = "msoSearchInMyNetworkPlaces"
        Case msoSearchInCustom: MsoSearchInToString = "msoSearchInCustom"
    End Select
End Function
