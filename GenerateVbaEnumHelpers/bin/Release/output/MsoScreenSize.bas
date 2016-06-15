Attribute VB_Name = "wMsoScreenSize"
Function MsoScreenSizeFromString(value As String) As MsoScreenSize
    If IsNumeric(value) Then
        MsoScreenSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoScreenSize544x376": MsoScreenSizeFromString = msoScreenSize544x376
        Case "msoScreenSize640x480": MsoScreenSizeFromString = msoScreenSize640x480
        Case "msoScreenSize720x512": MsoScreenSizeFromString = msoScreenSize720x512
        Case "msoScreenSize800x600": MsoScreenSizeFromString = msoScreenSize800x600
        Case "msoScreenSize1024x768": MsoScreenSizeFromString = msoScreenSize1024x768
        Case "msoScreenSize1152x882": MsoScreenSizeFromString = msoScreenSize1152x882
        Case "msoScreenSize1152x900": MsoScreenSizeFromString = msoScreenSize1152x900
        Case "msoScreenSize1280x1024": MsoScreenSizeFromString = msoScreenSize1280x1024
        Case "msoScreenSize1600x1200": MsoScreenSizeFromString = msoScreenSize1600x1200
        Case "msoScreenSize1800x1440": MsoScreenSizeFromString = msoScreenSize1800x1440
        Case "msoScreenSize1920x1200": MsoScreenSizeFromString = msoScreenSize1920x1200
    End Select
End Function

Function MsoScreenSizeToString(value As MsoScreenSize) As String
    Select Case value
        Case msoScreenSize544x376: MsoScreenSizeToString = "msoScreenSize544x376"
        Case msoScreenSize640x480: MsoScreenSizeToString = "msoScreenSize640x480"
        Case msoScreenSize720x512: MsoScreenSizeToString = "msoScreenSize720x512"
        Case msoScreenSize800x600: MsoScreenSizeToString = "msoScreenSize800x600"
        Case msoScreenSize1024x768: MsoScreenSizeToString = "msoScreenSize1024x768"
        Case msoScreenSize1152x882: MsoScreenSizeToString = "msoScreenSize1152x882"
        Case msoScreenSize1152x900: MsoScreenSizeToString = "msoScreenSize1152x900"
        Case msoScreenSize1280x1024: MsoScreenSizeToString = "msoScreenSize1280x1024"
        Case msoScreenSize1600x1200: MsoScreenSizeToString = "msoScreenSize1600x1200"
        Case msoScreenSize1800x1440: MsoScreenSizeToString = "msoScreenSize1800x1440"
        Case msoScreenSize1920x1200: MsoScreenSizeToString = "msoScreenSize1920x1200"
    End Select
End Function
