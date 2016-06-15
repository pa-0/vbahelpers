Attribute VB_Name = "wOlIconViewPlacement"
Function OlIconViewPlacementFromString(value As String) As OlIconViewPlacement
    If IsNumeric(value) Then
        OlIconViewPlacementFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olIconDoNotArrange": OlIconViewPlacementFromString = olIconDoNotArrange
        Case "olIconLineUp": OlIconViewPlacementFromString = olIconLineUp
        Case "olIconAutoArrange": OlIconViewPlacementFromString = olIconAutoArrange
        Case "olIconSortAndAutoArrange": OlIconViewPlacementFromString = olIconSortAndAutoArrange
    End Select
End Function

Function OlIconViewPlacementToString(value As OlIconViewPlacement) As String
    Select Case value
        Case olIconDoNotArrange: OlIconViewPlacementToString = "olIconDoNotArrange"
        Case olIconLineUp: OlIconViewPlacementToString = "olIconLineUp"
        Case olIconAutoArrange: OlIconViewPlacementToString = "olIconAutoArrange"
        Case olIconSortAndAutoArrange: OlIconViewPlacementToString = "olIconSortAndAutoArrange"
    End Select
End Function
