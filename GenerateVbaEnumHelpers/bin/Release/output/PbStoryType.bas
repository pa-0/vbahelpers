Attribute VB_Name = "wPbStoryType"
Function PbStoryTypeFromString(value As String) As PbStoryType
    If IsNumeric(value) Then
        PbStoryTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbStoryTable": PbStoryTypeFromString = pbStoryTable
        Case "pbStoryTextFrame": PbStoryTypeFromString = pbStoryTextFrame
        Case "pbStoryContinuedFrom": PbStoryTypeFromString = pbStoryContinuedFrom
        Case "pbStoryContinuedOn": PbStoryTypeFromString = pbStoryContinuedOn
    End Select
End Function

Function PbStoryTypeToString(value As PbStoryType) As String
    Select Case value
        Case pbStoryTable: PbStoryTypeToString = "pbStoryTable"
        Case pbStoryTextFrame: PbStoryTypeToString = "pbStoryTextFrame"
        Case pbStoryContinuedFrom: PbStoryTypeToString = "pbStoryContinuedFrom"
        Case pbStoryContinuedOn: PbStoryTypeToString = "pbStoryContinuedOn"
    End Select
End Function
