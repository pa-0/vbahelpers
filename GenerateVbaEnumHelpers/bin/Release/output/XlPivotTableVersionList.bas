Attribute VB_Name = "wXlPivotTableVersionList"
Function XlPivotTableVersionListFromString(value As String) As XlPivotTableVersionList
    If IsNumeric(value) Then
        XlPivotTableVersionListFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPivotTableVersion2000": XlPivotTableVersionListFromString = xlPivotTableVersion2000
        Case "xlPivotTableVersion10": XlPivotTableVersionListFromString = xlPivotTableVersion10
        Case "xlPivotTableVersion11": XlPivotTableVersionListFromString = xlPivotTableVersion11
        Case "xlPivotTableVersion12": XlPivotTableVersionListFromString = xlPivotTableVersion12
        Case "xlPivotTableVersion14": XlPivotTableVersionListFromString = xlPivotTableVersion14
        Case "xlPivotTableVersionCurrent": XlPivotTableVersionListFromString = xlPivotTableVersionCurrent
    End Select
End Function

Function XlPivotTableVersionListToString(value As XlPivotTableVersionList) As String
    Select Case value
        Case xlPivotTableVersion2000: XlPivotTableVersionListToString = "xlPivotTableVersion2000"
        Case xlPivotTableVersion10: XlPivotTableVersionListToString = "xlPivotTableVersion10"
        Case xlPivotTableVersion11: XlPivotTableVersionListToString = "xlPivotTableVersion11"
        Case xlPivotTableVersion12: XlPivotTableVersionListToString = "xlPivotTableVersion12"
        Case xlPivotTableVersion14: XlPivotTableVersionListToString = "xlPivotTableVersion14"
        Case xlPivotTableVersionCurrent: XlPivotTableVersionListToString = "xlPivotTableVersionCurrent"
    End Select
End Function
