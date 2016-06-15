Attribute VB_Name = "wMsoAppLanguageID"
Function MsoAppLanguageIDFromString(value As String) As MsoAppLanguageID
    If IsNumeric(value) Then
        MsoAppLanguageIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLanguageIDInstall": MsoAppLanguageIDFromString = msoLanguageIDInstall
        Case "msoLanguageIDUI": MsoAppLanguageIDFromString = msoLanguageIDUI
        Case "msoLanguageIDHelp": MsoAppLanguageIDFromString = msoLanguageIDHelp
        Case "msoLanguageIDExeMode": MsoAppLanguageIDFromString = msoLanguageIDExeMode
        Case "msoLanguageIDUIPrevious": MsoAppLanguageIDFromString = msoLanguageIDUIPrevious
    End Select
End Function

Function MsoAppLanguageIDToString(value As MsoAppLanguageID) As String
    Select Case value
        Case msoLanguageIDInstall: MsoAppLanguageIDToString = "msoLanguageIDInstall"
        Case msoLanguageIDUI: MsoAppLanguageIDToString = "msoLanguageIDUI"
        Case msoLanguageIDHelp: MsoAppLanguageIDToString = "msoLanguageIDHelp"
        Case msoLanguageIDExeMode: MsoAppLanguageIDToString = "msoLanguageIDExeMode"
        Case msoLanguageIDUIPrevious: MsoAppLanguageIDToString = "msoLanguageIDUIPrevious"
    End Select
End Function
