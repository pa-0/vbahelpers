Attribute VB_Name = "wMsoAnimProperty"
Function MsoAnimPropertyFromString(value As String) As MsoAnimProperty
    If IsNumeric(value) Then
        MsoAnimPropertyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimNone": MsoAnimPropertyFromString = msoAnimNone
        Case "msoAnimX": MsoAnimPropertyFromString = msoAnimX
        Case "msoAnimY": MsoAnimPropertyFromString = msoAnimY
        Case "msoAnimWidth": MsoAnimPropertyFromString = msoAnimWidth
        Case "msoAnimHeight": MsoAnimPropertyFromString = msoAnimHeight
        Case "msoAnimOpacity": MsoAnimPropertyFromString = msoAnimOpacity
        Case "msoAnimRotation": MsoAnimPropertyFromString = msoAnimRotation
        Case "msoAnimColor": MsoAnimPropertyFromString = msoAnimColor
        Case "msoAnimVisibility": MsoAnimPropertyFromString = msoAnimVisibility
        Case "msoAnimTextFontBold": MsoAnimPropertyFromString = msoAnimTextFontBold
        Case "msoAnimTextFontColor": MsoAnimPropertyFromString = msoAnimTextFontColor
        Case "msoAnimTextFontEmboss": MsoAnimPropertyFromString = msoAnimTextFontEmboss
        Case "msoAnimTextFontItalic": MsoAnimPropertyFromString = msoAnimTextFontItalic
        Case "msoAnimTextFontName": MsoAnimPropertyFromString = msoAnimTextFontName
        Case "msoAnimTextFontShadow": MsoAnimPropertyFromString = msoAnimTextFontShadow
        Case "msoAnimTextFontSize": MsoAnimPropertyFromString = msoAnimTextFontSize
        Case "msoAnimTextFontSubscript": MsoAnimPropertyFromString = msoAnimTextFontSubscript
        Case "msoAnimTextFontSuperscript": MsoAnimPropertyFromString = msoAnimTextFontSuperscript
        Case "msoAnimTextFontUnderline": MsoAnimPropertyFromString = msoAnimTextFontUnderline
        Case "msoAnimTextFontStrikeThrough": MsoAnimPropertyFromString = msoAnimTextFontStrikeThrough
        Case "msoAnimTextBulletCharacter": MsoAnimPropertyFromString = msoAnimTextBulletCharacter
        Case "msoAnimTextBulletFontName": MsoAnimPropertyFromString = msoAnimTextBulletFontName
        Case "msoAnimTextBulletNumber": MsoAnimPropertyFromString = msoAnimTextBulletNumber
        Case "msoAnimTextBulletColor": MsoAnimPropertyFromString = msoAnimTextBulletColor
        Case "msoAnimTextBulletRelativeSize": MsoAnimPropertyFromString = msoAnimTextBulletRelativeSize
        Case "msoAnimTextBulletStyle": MsoAnimPropertyFromString = msoAnimTextBulletStyle
        Case "msoAnimTextBulletType": MsoAnimPropertyFromString = msoAnimTextBulletType
        Case "msoAnimShapePictureContrast": MsoAnimPropertyFromString = msoAnimShapePictureContrast
        Case "msoAnimShapePictureBrightness": MsoAnimPropertyFromString = msoAnimShapePictureBrightness
        Case "msoAnimShapePictureGamma": MsoAnimPropertyFromString = msoAnimShapePictureGamma
        Case "msoAnimShapePictureGrayscale": MsoAnimPropertyFromString = msoAnimShapePictureGrayscale
        Case "msoAnimShapeFillOn": MsoAnimPropertyFromString = msoAnimShapeFillOn
        Case "msoAnimShapeFillColor": MsoAnimPropertyFromString = msoAnimShapeFillColor
        Case "msoAnimShapeFillOpacity": MsoAnimPropertyFromString = msoAnimShapeFillOpacity
        Case "msoAnimShapeFillBackColor": MsoAnimPropertyFromString = msoAnimShapeFillBackColor
        Case "msoAnimShapeLineOn": MsoAnimPropertyFromString = msoAnimShapeLineOn
        Case "msoAnimShapeLineColor": MsoAnimPropertyFromString = msoAnimShapeLineColor
        Case "msoAnimShapeShadowOn": MsoAnimPropertyFromString = msoAnimShapeShadowOn
        Case "msoAnimShapeShadowType": MsoAnimPropertyFromString = msoAnimShapeShadowType
        Case "msoAnimShapeShadowColor": MsoAnimPropertyFromString = msoAnimShapeShadowColor
        Case "msoAnimShapeShadowOpacity": MsoAnimPropertyFromString = msoAnimShapeShadowOpacity
        Case "msoAnimShapeShadowOffsetX": MsoAnimPropertyFromString = msoAnimShapeShadowOffsetX
        Case "msoAnimShapeShadowOffsetY": MsoAnimPropertyFromString = msoAnimShapeShadowOffsetY
    End Select
End Function

Function MsoAnimPropertyToString(value As MsoAnimProperty) As String
    Select Case value
        Case msoAnimNone: MsoAnimPropertyToString = "msoAnimNone"
        Case msoAnimX: MsoAnimPropertyToString = "msoAnimX"
        Case msoAnimY: MsoAnimPropertyToString = "msoAnimY"
        Case msoAnimWidth: MsoAnimPropertyToString = "msoAnimWidth"
        Case msoAnimHeight: MsoAnimPropertyToString = "msoAnimHeight"
        Case msoAnimOpacity: MsoAnimPropertyToString = "msoAnimOpacity"
        Case msoAnimRotation: MsoAnimPropertyToString = "msoAnimRotation"
        Case msoAnimColor: MsoAnimPropertyToString = "msoAnimColor"
        Case msoAnimVisibility: MsoAnimPropertyToString = "msoAnimVisibility"
        Case msoAnimTextFontBold: MsoAnimPropertyToString = "msoAnimTextFontBold"
        Case msoAnimTextFontColor: MsoAnimPropertyToString = "msoAnimTextFontColor"
        Case msoAnimTextFontEmboss: MsoAnimPropertyToString = "msoAnimTextFontEmboss"
        Case msoAnimTextFontItalic: MsoAnimPropertyToString = "msoAnimTextFontItalic"
        Case msoAnimTextFontName: MsoAnimPropertyToString = "msoAnimTextFontName"
        Case msoAnimTextFontShadow: MsoAnimPropertyToString = "msoAnimTextFontShadow"
        Case msoAnimTextFontSize: MsoAnimPropertyToString = "msoAnimTextFontSize"
        Case msoAnimTextFontSubscript: MsoAnimPropertyToString = "msoAnimTextFontSubscript"
        Case msoAnimTextFontSuperscript: MsoAnimPropertyToString = "msoAnimTextFontSuperscript"
        Case msoAnimTextFontUnderline: MsoAnimPropertyToString = "msoAnimTextFontUnderline"
        Case msoAnimTextFontStrikeThrough: MsoAnimPropertyToString = "msoAnimTextFontStrikeThrough"
        Case msoAnimTextBulletCharacter: MsoAnimPropertyToString = "msoAnimTextBulletCharacter"
        Case msoAnimTextBulletFontName: MsoAnimPropertyToString = "msoAnimTextBulletFontName"
        Case msoAnimTextBulletNumber: MsoAnimPropertyToString = "msoAnimTextBulletNumber"
        Case msoAnimTextBulletColor: MsoAnimPropertyToString = "msoAnimTextBulletColor"
        Case msoAnimTextBulletRelativeSize: MsoAnimPropertyToString = "msoAnimTextBulletRelativeSize"
        Case msoAnimTextBulletStyle: MsoAnimPropertyToString = "msoAnimTextBulletStyle"
        Case msoAnimTextBulletType: MsoAnimPropertyToString = "msoAnimTextBulletType"
        Case msoAnimShapePictureContrast: MsoAnimPropertyToString = "msoAnimShapePictureContrast"
        Case msoAnimShapePictureBrightness: MsoAnimPropertyToString = "msoAnimShapePictureBrightness"
        Case msoAnimShapePictureGamma: MsoAnimPropertyToString = "msoAnimShapePictureGamma"
        Case msoAnimShapePictureGrayscale: MsoAnimPropertyToString = "msoAnimShapePictureGrayscale"
        Case msoAnimShapeFillOn: MsoAnimPropertyToString = "msoAnimShapeFillOn"
        Case msoAnimShapeFillColor: MsoAnimPropertyToString = "msoAnimShapeFillColor"
        Case msoAnimShapeFillOpacity: MsoAnimPropertyToString = "msoAnimShapeFillOpacity"
        Case msoAnimShapeFillBackColor: MsoAnimPropertyToString = "msoAnimShapeFillBackColor"
        Case msoAnimShapeLineOn: MsoAnimPropertyToString = "msoAnimShapeLineOn"
        Case msoAnimShapeLineColor: MsoAnimPropertyToString = "msoAnimShapeLineColor"
        Case msoAnimShapeShadowOn: MsoAnimPropertyToString = "msoAnimShapeShadowOn"
        Case msoAnimShapeShadowType: MsoAnimPropertyToString = "msoAnimShapeShadowType"
        Case msoAnimShapeShadowColor: MsoAnimPropertyToString = "msoAnimShapeShadowColor"
        Case msoAnimShapeShadowOpacity: MsoAnimPropertyToString = "msoAnimShapeShadowOpacity"
        Case msoAnimShapeShadowOffsetX: MsoAnimPropertyToString = "msoAnimShapeShadowOffsetX"
        Case msoAnimShapeShadowOffsetY: MsoAnimPropertyToString = "msoAnimShapeShadowOffsetY"
    End Select
End Function
