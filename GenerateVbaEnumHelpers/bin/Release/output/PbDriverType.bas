Attribute VB_Name = "wPbDriverType"
Function PbDriverTypeFromString(value As String) As PbDriverType
    If IsNumeric(value) Then
        PbDriverTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbDriverTypeNonPostScript": PbDriverTypeFromString = pbDriverTypeNonPostScript
        Case "pbDriverTypePostScript1": PbDriverTypeFromString = pbDriverTypePostScript1
        Case "pbDriverTypePostScript2": PbDriverTypeFromString = pbDriverTypePostScript2
        Case "pbDriverTypePostScript3": PbDriverTypeFromString = pbDriverTypePostScript3
        Case "pbDriverTypeXPS": PbDriverTypeFromString = pbDriverTypeXPS
    End Select
End Function

Function PbDriverTypeToString(value As PbDriverType) As String
    Select Case value
        Case pbDriverTypeNonPostScript: PbDriverTypeToString = "pbDriverTypeNonPostScript"
        Case pbDriverTypePostScript1: PbDriverTypeToString = "pbDriverTypePostScript1"
        Case pbDriverTypePostScript2: PbDriverTypeToString = "pbDriverTypePostScript2"
        Case pbDriverTypePostScript3: PbDriverTypeToString = "pbDriverTypePostScript3"
        Case pbDriverTypeXPS: PbDriverTypeToString = "pbDriverTypeXPS"
    End Select
End Function
