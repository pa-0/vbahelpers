Attribute VB_Name = "wXlDeleteShiftDirection"
Function XlDeleteShiftDirectionFromString(value As String) As XlDeleteShiftDirection
    If IsNumeric(value) Then
        XlDeleteShiftDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlShiftUp": XlDeleteShiftDirectionFromString = xlShiftUp
        Case "xlShiftToLeft": XlDeleteShiftDirectionFromString = xlShiftToLeft
    End Select
End Function

Function XlDeleteShiftDirectionToString(value As XlDeleteShiftDirection) As String
    Select Case value
        Case xlShiftUp: XlDeleteShiftDirectionToString = "xlShiftUp"
        Case xlShiftToLeft: XlDeleteShiftDirectionToString = "xlShiftToLeft"
    End Select
End Function
