Attribute VB_Name = "wMsoAnimTriggerType"
Function MsoAnimTriggerTypeFromString(value As String) As MsoAnimTriggerType
    If IsNumeric(value) Then
        MsoAnimTriggerTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimTriggerNone": MsoAnimTriggerTypeFromString = msoAnimTriggerNone
        Case "msoAnimTriggerOnPageClick": MsoAnimTriggerTypeFromString = msoAnimTriggerOnPageClick
        Case "msoAnimTriggerWithPrevious": MsoAnimTriggerTypeFromString = msoAnimTriggerWithPrevious
        Case "msoAnimTriggerAfterPrevious": MsoAnimTriggerTypeFromString = msoAnimTriggerAfterPrevious
        Case "msoAnimTriggerOnShapeClick": MsoAnimTriggerTypeFromString = msoAnimTriggerOnShapeClick
        Case "msoAnimTriggerOnMediaBookmark": MsoAnimTriggerTypeFromString = msoAnimTriggerOnMediaBookmark
        Case "msoAnimTriggerMixed": MsoAnimTriggerTypeFromString = msoAnimTriggerMixed
    End Select
End Function

Function MsoAnimTriggerTypeToString(value As MsoAnimTriggerType) As String
    Select Case value
        Case msoAnimTriggerNone: MsoAnimTriggerTypeToString = "msoAnimTriggerNone"
        Case msoAnimTriggerOnPageClick: MsoAnimTriggerTypeToString = "msoAnimTriggerOnPageClick"
        Case msoAnimTriggerWithPrevious: MsoAnimTriggerTypeToString = "msoAnimTriggerWithPrevious"
        Case msoAnimTriggerAfterPrevious: MsoAnimTriggerTypeToString = "msoAnimTriggerAfterPrevious"
        Case msoAnimTriggerOnShapeClick: MsoAnimTriggerTypeToString = "msoAnimTriggerOnShapeClick"
        Case msoAnimTriggerOnMediaBookmark: MsoAnimTriggerTypeToString = "msoAnimTriggerOnMediaBookmark"
        Case msoAnimTriggerMixed: MsoAnimTriggerTypeToString = "msoAnimTriggerMixed"
    End Select
End Function
