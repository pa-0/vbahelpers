Attribute VB_Name = "wPpPlaceholderType"
Function PpPlaceholderTypeFromString(value As String) As PpPlaceholderType
    If IsNumeric(value) Then
        PpPlaceholderTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPlaceholderTitle": PpPlaceholderTypeFromString = ppPlaceholderTitle
        Case "ppPlaceholderBody": PpPlaceholderTypeFromString = ppPlaceholderBody
        Case "ppPlaceholderCenterTitle": PpPlaceholderTypeFromString = ppPlaceholderCenterTitle
        Case "ppPlaceholderSubtitle": PpPlaceholderTypeFromString = ppPlaceholderSubtitle
        Case "ppPlaceholderVerticalTitle": PpPlaceholderTypeFromString = ppPlaceholderVerticalTitle
        Case "ppPlaceholderVerticalBody": PpPlaceholderTypeFromString = ppPlaceholderVerticalBody
        Case "ppPlaceholderObject": PpPlaceholderTypeFromString = ppPlaceholderObject
        Case "ppPlaceholderChart": PpPlaceholderTypeFromString = ppPlaceholderChart
        Case "ppPlaceholderBitmap": PpPlaceholderTypeFromString = ppPlaceholderBitmap
        Case "ppPlaceholderMediaClip": PpPlaceholderTypeFromString = ppPlaceholderMediaClip
        Case "ppPlaceholderOrgChart": PpPlaceholderTypeFromString = ppPlaceholderOrgChart
        Case "ppPlaceholderTable": PpPlaceholderTypeFromString = ppPlaceholderTable
        Case "ppPlaceholderSlideNumber": PpPlaceholderTypeFromString = ppPlaceholderSlideNumber
        Case "ppPlaceholderHeader": PpPlaceholderTypeFromString = ppPlaceholderHeader
        Case "ppPlaceholderFooter": PpPlaceholderTypeFromString = ppPlaceholderFooter
        Case "ppPlaceholderDate": PpPlaceholderTypeFromString = ppPlaceholderDate
        Case "ppPlaceholderVerticalObject": PpPlaceholderTypeFromString = ppPlaceholderVerticalObject
        Case "ppPlaceholderPicture": PpPlaceholderTypeFromString = ppPlaceholderPicture
        Case "ppPlaceholderMixed": PpPlaceholderTypeFromString = ppPlaceholderMixed
    End Select
End Function

Function PpPlaceholderTypeToString(value As PpPlaceholderType) As String
    Select Case value
        Case ppPlaceholderTitle: PpPlaceholderTypeToString = "ppPlaceholderTitle"
        Case ppPlaceholderBody: PpPlaceholderTypeToString = "ppPlaceholderBody"
        Case ppPlaceholderCenterTitle: PpPlaceholderTypeToString = "ppPlaceholderCenterTitle"
        Case ppPlaceholderSubtitle: PpPlaceholderTypeToString = "ppPlaceholderSubtitle"
        Case ppPlaceholderVerticalTitle: PpPlaceholderTypeToString = "ppPlaceholderVerticalTitle"
        Case ppPlaceholderVerticalBody: PpPlaceholderTypeToString = "ppPlaceholderVerticalBody"
        Case ppPlaceholderObject: PpPlaceholderTypeToString = "ppPlaceholderObject"
        Case ppPlaceholderChart: PpPlaceholderTypeToString = "ppPlaceholderChart"
        Case ppPlaceholderBitmap: PpPlaceholderTypeToString = "ppPlaceholderBitmap"
        Case ppPlaceholderMediaClip: PpPlaceholderTypeToString = "ppPlaceholderMediaClip"
        Case ppPlaceholderOrgChart: PpPlaceholderTypeToString = "ppPlaceholderOrgChart"
        Case ppPlaceholderTable: PpPlaceholderTypeToString = "ppPlaceholderTable"
        Case ppPlaceholderSlideNumber: PpPlaceholderTypeToString = "ppPlaceholderSlideNumber"
        Case ppPlaceholderHeader: PpPlaceholderTypeToString = "ppPlaceholderHeader"
        Case ppPlaceholderFooter: PpPlaceholderTypeToString = "ppPlaceholderFooter"
        Case ppPlaceholderDate: PpPlaceholderTypeToString = "ppPlaceholderDate"
        Case ppPlaceholderVerticalObject: PpPlaceholderTypeToString = "ppPlaceholderVerticalObject"
        Case ppPlaceholderPicture: PpPlaceholderTypeToString = "ppPlaceholderPicture"
        Case ppPlaceholderMixed: PpPlaceholderTypeToString = "ppPlaceholderMixed"
    End Select
End Function
