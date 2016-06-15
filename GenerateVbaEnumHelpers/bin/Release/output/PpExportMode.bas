Attribute VB_Name = "wPpExportMode"
Function PpExportModeFromString(value As String) As PpExportMode
    If IsNumeric(value) Then
        PpExportModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppRelativeToSlide": PpExportModeFromString = ppRelativeToSlide
        Case "ppClipRelativeToSlide": PpExportModeFromString = ppClipRelativeToSlide
        Case "ppScaleToFit": PpExportModeFromString = ppScaleToFit
        Case "ppScaleXY": PpExportModeFromString = ppScaleXY
    End Select
End Function

Function PpExportModeToString(value As PpExportMode) As String
    Select Case value
        Case ppRelativeToSlide: PpExportModeToString = "ppRelativeToSlide"
        Case ppClipRelativeToSlide: PpExportModeToString = "ppClipRelativeToSlide"
        Case ppScaleToFit: PpExportModeToString = "ppScaleToFit"
        Case ppScaleXY: PpExportModeToString = "ppScaleXY"
    End Select
End Function
