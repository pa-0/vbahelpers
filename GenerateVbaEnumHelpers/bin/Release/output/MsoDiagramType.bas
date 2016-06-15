Attribute VB_Name = "wMsoDiagramType"
Function MsoDiagramTypeFromString(value As String) As MsoDiagramType
    If IsNumeric(value) Then
        MsoDiagramTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoDiagramOrgChart": MsoDiagramTypeFromString = msoDiagramOrgChart
        Case "msoDiagramCycle": MsoDiagramTypeFromString = msoDiagramCycle
        Case "msoDiagramRadial": MsoDiagramTypeFromString = msoDiagramRadial
        Case "msoDiagramPyramid": MsoDiagramTypeFromString = msoDiagramPyramid
        Case "msoDiagramVenn": MsoDiagramTypeFromString = msoDiagramVenn
        Case "msoDiagramTarget": MsoDiagramTypeFromString = msoDiagramTarget
        Case "msoDiagramMixed": MsoDiagramTypeFromString = msoDiagramMixed
    End Select
End Function

Function MsoDiagramTypeToString(value As MsoDiagramType) As String
    Select Case value
        Case msoDiagramOrgChart: MsoDiagramTypeToString = "msoDiagramOrgChart"
        Case msoDiagramCycle: MsoDiagramTypeToString = "msoDiagramCycle"
        Case msoDiagramRadial: MsoDiagramTypeToString = "msoDiagramRadial"
        Case msoDiagramPyramid: MsoDiagramTypeToString = "msoDiagramPyramid"
        Case msoDiagramVenn: MsoDiagramTypeToString = "msoDiagramVenn"
        Case msoDiagramTarget: MsoDiagramTypeToString = "msoDiagramTarget"
        Case msoDiagramMixed: MsoDiagramTypeToString = "msoDiagramMixed"
    End Select
End Function
