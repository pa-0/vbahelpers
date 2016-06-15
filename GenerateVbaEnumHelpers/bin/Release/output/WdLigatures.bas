Attribute VB_Name = "wWdLigatures"
Function WdLigaturesFromString(value As String) As WdLigatures
    If IsNumeric(value) Then
        WdLigaturesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLigaturesNone": WdLigaturesFromString = wdLigaturesNone
        Case "wdLigaturesStandard": WdLigaturesFromString = wdLigaturesStandard
        Case "wdLigaturesContextual": WdLigaturesFromString = wdLigaturesContextual
        Case "wdLigaturesStandardContextual": WdLigaturesFromString = wdLigaturesStandardContextual
        Case "wdLigaturesHistorical": WdLigaturesFromString = wdLigaturesHistorical
        Case "wdLigaturesStandardHistorical": WdLigaturesFromString = wdLigaturesStandardHistorical
        Case "wdLigaturesContextualHistorical": WdLigaturesFromString = wdLigaturesContextualHistorical
        Case "wdLigaturesStandardContextualHistorical": WdLigaturesFromString = wdLigaturesStandardContextualHistorical
        Case "wdLigaturesDiscretional": WdLigaturesFromString = wdLigaturesDiscretional
        Case "wdLigaturesStandardDiscretional": WdLigaturesFromString = wdLigaturesStandardDiscretional
        Case "wdLigaturesContextualDiscretional": WdLigaturesFromString = wdLigaturesContextualDiscretional
        Case "wdLigaturesStandardContextualDiscretional": WdLigaturesFromString = wdLigaturesStandardContextualDiscretional
        Case "wdLigaturesHistoricalDiscretional": WdLigaturesFromString = wdLigaturesHistoricalDiscretional
        Case "wdLigaturesStandardHistoricalDiscretional": WdLigaturesFromString = wdLigaturesStandardHistoricalDiscretional
        Case "wdLigaturesContextualHistoricalDiscretional": WdLigaturesFromString = wdLigaturesContextualHistoricalDiscretional
        Case "wdLigaturesAll": WdLigaturesFromString = wdLigaturesAll
    End Select
End Function

Function WdLigaturesToString(value As WdLigatures) As String
    Select Case value
        Case wdLigaturesNone: WdLigaturesToString = "wdLigaturesNone"
        Case wdLigaturesStandard: WdLigaturesToString = "wdLigaturesStandard"
        Case wdLigaturesContextual: WdLigaturesToString = "wdLigaturesContextual"
        Case wdLigaturesStandardContextual: WdLigaturesToString = "wdLigaturesStandardContextual"
        Case wdLigaturesHistorical: WdLigaturesToString = "wdLigaturesHistorical"
        Case wdLigaturesStandardHistorical: WdLigaturesToString = "wdLigaturesStandardHistorical"
        Case wdLigaturesContextualHistorical: WdLigaturesToString = "wdLigaturesContextualHistorical"
        Case wdLigaturesStandardContextualHistorical: WdLigaturesToString = "wdLigaturesStandardContextualHistorical"
        Case wdLigaturesDiscretional: WdLigaturesToString = "wdLigaturesDiscretional"
        Case wdLigaturesStandardDiscretional: WdLigaturesToString = "wdLigaturesStandardDiscretional"
        Case wdLigaturesContextualDiscretional: WdLigaturesToString = "wdLigaturesContextualDiscretional"
        Case wdLigaturesStandardContextualDiscretional: WdLigaturesToString = "wdLigaturesStandardContextualDiscretional"
        Case wdLigaturesHistoricalDiscretional: WdLigaturesToString = "wdLigaturesHistoricalDiscretional"
        Case wdLigaturesStandardHistoricalDiscretional: WdLigaturesToString = "wdLigaturesStandardHistoricalDiscretional"
        Case wdLigaturesContextualHistoricalDiscretional: WdLigaturesToString = "wdLigaturesContextualHistoricalDiscretional"
        Case wdLigaturesAll: WdLigaturesToString = "wdLigaturesAll"
    End Select
End Function
