Attribute VB_Name = "wMsoCalloutType"
Function MsoCalloutTypeFromString(value As String) As MsoCalloutType
    If IsNumeric(value) Then
        MsoCalloutTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCalloutOne": MsoCalloutTypeFromString = msoCalloutOne
        Case "msoCalloutTwo": MsoCalloutTypeFromString = msoCalloutTwo
        Case "msoCalloutThree": MsoCalloutTypeFromString = msoCalloutThree
        Case "msoCalloutFour": MsoCalloutTypeFromString = msoCalloutFour
        Case "msoCalloutMixed": MsoCalloutTypeFromString = msoCalloutMixed
    End Select
End Function

Function MsoCalloutTypeToString(value As MsoCalloutType) As String
    Select Case value
        Case msoCalloutOne: MsoCalloutTypeToString = "msoCalloutOne"
        Case msoCalloutTwo: MsoCalloutTypeToString = "msoCalloutTwo"
        Case msoCalloutThree: MsoCalloutTypeToString = "msoCalloutThree"
        Case msoCalloutFour: MsoCalloutTypeToString = "msoCalloutFour"
        Case msoCalloutMixed: MsoCalloutTypeToString = "msoCalloutMixed"
    End Select
End Function
