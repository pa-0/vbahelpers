Attribute VB_Name = "wPpMouseActivation"
Function PpMouseActivationFromString(value As String) As PpMouseActivation
    If IsNumeric(value) Then
        PpMouseActivationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppMouseClick": PpMouseActivationFromString = ppMouseClick
        Case "ppMouseOver": PpMouseActivationFromString = ppMouseOver
    End Select
End Function

Function PpMouseActivationToString(value As PpMouseActivation) As String
    Select Case value
        Case ppMouseClick: PpMouseActivationToString = "ppMouseClick"
        Case ppMouseOver: PpMouseActivationToString = "ppMouseOver"
    End Select
End Function
