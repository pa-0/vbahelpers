Attribute VB_Name = "wMsoScriptLocation"
Function MsoScriptLocationFromString(value As String) As MsoScriptLocation
    If IsNumeric(value) Then
        MsoScriptLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoScriptLocationInHead": MsoScriptLocationFromString = msoScriptLocationInHead
        Case "msoScriptLocationInBody": MsoScriptLocationFromString = msoScriptLocationInBody
    End Select
End Function

Function MsoScriptLocationToString(value As MsoScriptLocation) As String
    Select Case value
        Case msoScriptLocationInHead: MsoScriptLocationToString = "msoScriptLocationInHead"
        Case msoScriptLocationInBody: MsoScriptLocationToString = "msoScriptLocationInBody"
    End Select
End Function
