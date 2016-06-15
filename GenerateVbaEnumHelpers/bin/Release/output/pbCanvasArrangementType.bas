Attribute VB_Name = "wpbCanvasArrangementType"
Function pbCanvasArrangementTypeFromString(value As String) As pbCanvasArrangementType
    If IsNumeric(value) Then
        pbCanvasArrangementTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbCanvasArrangementTypeOneCanvas": pbCanvasArrangementTypeFromString = pbCanvasArrangementTypeOneCanvas
        Case "pbCanvasArrangementTypeColsCanvas": pbCanvasArrangementTypeFromString = pbCanvasArrangementTypeColsCanvas
        Case "pbCanvasArrangementTypeRowsCanvas": pbCanvasArrangementTypeFromString = pbCanvasArrangementTypeRowsCanvas
    End Select
End Function

Function pbCanvasArrangementTypeToString(value As pbCanvasArrangementType) As String
    Select Case value
        Case pbCanvasArrangementTypeOneCanvas: pbCanvasArrangementTypeToString = "pbCanvasArrangementTypeOneCanvas"
        Case pbCanvasArrangementTypeColsCanvas: pbCanvasArrangementTypeToString = "pbCanvasArrangementTypeColsCanvas"
        Case pbCanvasArrangementTypeRowsCanvas: pbCanvasArrangementTypeToString = "pbCanvasArrangementTypeRowsCanvas"
    End Select
End Function
