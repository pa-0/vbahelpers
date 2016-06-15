Attribute VB_Name = "wXlMSApplication"
Function XlMSApplicationFromString(value As String) As XlMSApplication
    If IsNumeric(value) Then
        XlMSApplicationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlMicrosoftWord": XlMSApplicationFromString = xlMicrosoftWord
        Case "xlMicrosoftPowerPoint": XlMSApplicationFromString = xlMicrosoftPowerPoint
        Case "xlMicrosoftMail": XlMSApplicationFromString = xlMicrosoftMail
        Case "xlMicrosoftAccess": XlMSApplicationFromString = xlMicrosoftAccess
        Case "xlMicrosoftFoxPro": XlMSApplicationFromString = xlMicrosoftFoxPro
        Case "xlMicrosoftProject": XlMSApplicationFromString = xlMicrosoftProject
        Case "xlMicrosoftSchedulePlus": XlMSApplicationFromString = xlMicrosoftSchedulePlus
    End Select
End Function

Function XlMSApplicationToString(value As XlMSApplication) As String
    Select Case value
        Case xlMicrosoftWord: XlMSApplicationToString = "xlMicrosoftWord"
        Case xlMicrosoftPowerPoint: XlMSApplicationToString = "xlMicrosoftPowerPoint"
        Case xlMicrosoftMail: XlMSApplicationToString = "xlMicrosoftMail"
        Case xlMicrosoftAccess: XlMSApplicationToString = "xlMicrosoftAccess"
        Case xlMicrosoftFoxPro: XlMSApplicationToString = "xlMicrosoftFoxPro"
        Case xlMicrosoftProject: XlMSApplicationToString = "xlMicrosoftProject"
        Case xlMicrosoftSchedulePlus: XlMSApplicationToString = "xlMicrosoftSchedulePlus"
    End Select
End Function
