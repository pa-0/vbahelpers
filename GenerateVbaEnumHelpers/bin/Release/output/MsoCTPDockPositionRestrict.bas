Attribute VB_Name = "wMsoCTPDockPositionRestrict"
Function MsoCTPDockPositionRestrictFromString(value As String) As MsoCTPDockPositionRestrict
    If IsNumeric(value) Then
        MsoCTPDockPositionRestrictFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCTPDockPositionRestrictNone": MsoCTPDockPositionRestrictFromString = msoCTPDockPositionRestrictNone
        Case "msoCTPDockPositionRestrictNoChange": MsoCTPDockPositionRestrictFromString = msoCTPDockPositionRestrictNoChange
        Case "msoCTPDockPositionRestrictNoHorizontal": MsoCTPDockPositionRestrictFromString = msoCTPDockPositionRestrictNoHorizontal
        Case "msoCTPDockPositionRestrictNoVertical": MsoCTPDockPositionRestrictFromString = msoCTPDockPositionRestrictNoVertical
    End Select
End Function

Function MsoCTPDockPositionRestrictToString(value As MsoCTPDockPositionRestrict) As String
    Select Case value
        Case msoCTPDockPositionRestrictNone: MsoCTPDockPositionRestrictToString = "msoCTPDockPositionRestrictNone"
        Case msoCTPDockPositionRestrictNoChange: MsoCTPDockPositionRestrictToString = "msoCTPDockPositionRestrictNoChange"
        Case msoCTPDockPositionRestrictNoHorizontal: MsoCTPDockPositionRestrictToString = "msoCTPDockPositionRestrictNoHorizontal"
        Case msoCTPDockPositionRestrictNoVertical: MsoCTPDockPositionRestrictToString = "msoCTPDockPositionRestrictNoVertical"
    End Select
End Function
