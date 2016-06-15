Attribute VB_Name = "wXlInsertShiftDirection"
Function XlInsertShiftDirectionFromString(value As String) As XlInsertShiftDirection
    If IsNumeric(value) Then
        XlInsertShiftDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlShiftToRight": XlInsertShiftDirectionFromString = xlShiftToRight
        Case "xlShiftDown": XlInsertShiftDirectionFromString = xlShiftDown
    End Select
End Function

Function XlInsertShiftDirectionToString(value As XlInsertShiftDirection) As String
    Select Case value
        Case xlShiftToRight: XlInsertShiftDirectionToString = "xlShiftToRight"
        Case xlShiftDown: XlInsertShiftDirectionToString = "xlShiftDown"
    End Select
End Function
