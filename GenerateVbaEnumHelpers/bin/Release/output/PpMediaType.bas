Attribute VB_Name = "wPpMediaType"
Function PpMediaTypeFromString(value As String) As PpMediaType
    If IsNumeric(value) Then
        PpMediaTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppMediaTypeOther": PpMediaTypeFromString = ppMediaTypeOther
        Case "ppMediaTypeSound": PpMediaTypeFromString = ppMediaTypeSound
        Case "ppMediaTypeMovie": PpMediaTypeFromString = ppMediaTypeMovie
        Case "ppMediaTypeMixed": PpMediaTypeFromString = ppMediaTypeMixed
    End Select
End Function

Function PpMediaTypeToString(value As PpMediaType) As String
    Select Case value
        Case ppMediaTypeOther: PpMediaTypeToString = "ppMediaTypeOther"
        Case ppMediaTypeSound: PpMediaTypeToString = "ppMediaTypeSound"
        Case ppMediaTypeMovie: PpMediaTypeToString = "ppMediaTypeMovie"
        Case ppMediaTypeMixed: PpMediaTypeToString = "ppMediaTypeMixed"
    End Select
End Function
