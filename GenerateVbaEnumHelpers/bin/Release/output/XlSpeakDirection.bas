Attribute VB_Name = "wXlSpeakDirection"
Function XlSpeakDirectionFromString(value As String) As XlSpeakDirection
    If IsNumeric(value) Then
        XlSpeakDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSpeakByRows": XlSpeakDirectionFromString = xlSpeakByRows
        Case "xlSpeakByColumns": XlSpeakDirectionFromString = xlSpeakByColumns
    End Select
End Function

Function XlSpeakDirectionToString(value As XlSpeakDirection) As String
    Select Case value
        Case xlSpeakByRows: XlSpeakDirectionToString = "xlSpeakByRows"
        Case xlSpeakByColumns: XlSpeakDirectionToString = "xlSpeakByColumns"
    End Select
End Function
