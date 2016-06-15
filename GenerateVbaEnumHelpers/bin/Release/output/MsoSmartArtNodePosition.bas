Attribute VB_Name = "wMsoSmartArtNodePosition"
Function MsoSmartArtNodePositionFromString(value As String) As MsoSmartArtNodePosition
    If IsNumeric(value) Then
        MsoSmartArtNodePositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSmartArtNodeDefault": MsoSmartArtNodePositionFromString = msoSmartArtNodeDefault
        Case "msoSmartArtNodeAfter": MsoSmartArtNodePositionFromString = msoSmartArtNodeAfter
        Case "msoSmartArtNodeBefore": MsoSmartArtNodePositionFromString = msoSmartArtNodeBefore
        Case "msoSmartArtNodeAbove": MsoSmartArtNodePositionFromString = msoSmartArtNodeAbove
        Case "msoSmartArtNodeBelow": MsoSmartArtNodePositionFromString = msoSmartArtNodeBelow
    End Select
End Function

Function MsoSmartArtNodePositionToString(value As MsoSmartArtNodePosition) As String
    Select Case value
        Case msoSmartArtNodeDefault: MsoSmartArtNodePositionToString = "msoSmartArtNodeDefault"
        Case msoSmartArtNodeAfter: MsoSmartArtNodePositionToString = "msoSmartArtNodeAfter"
        Case msoSmartArtNodeBefore: MsoSmartArtNodePositionToString = "msoSmartArtNodeBefore"
        Case msoSmartArtNodeAbove: MsoSmartArtNodePositionToString = "msoSmartArtNodeAbove"
        Case msoSmartArtNodeBelow: MsoSmartArtNodePositionToString = "msoSmartArtNodeBelow"
    End Select
End Function
