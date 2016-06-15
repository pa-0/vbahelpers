Attribute VB_Name = "wRibbonControlSize"
Function RibbonControlSizeFromString(value As String) As RibbonControlSize
    If IsNumeric(value) Then
        RibbonControlSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "RibbonControlSizeRegular": RibbonControlSizeFromString = RibbonControlSizeRegular
        Case "RibbonControlSizeLarge": RibbonControlSizeFromString = RibbonControlSizeLarge
    End Select
End Function

Function RibbonControlSizeToString(value As RibbonControlSize) As String
    Select Case value
        Case RibbonControlSizeRegular: RibbonControlSizeToString = "RibbonControlSizeRegular"
        Case RibbonControlSizeLarge: RibbonControlSizeToString = "RibbonControlSizeLarge"
    End Select
End Function
