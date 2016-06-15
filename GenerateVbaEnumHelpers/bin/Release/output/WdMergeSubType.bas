Attribute VB_Name = "wWdMergeSubType"
Function WdMergeSubTypeFromString(value As String) As WdMergeSubType
    If IsNumeric(value) Then
        WdMergeSubTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMergeSubTypeOther": WdMergeSubTypeFromString = wdMergeSubTypeOther
        Case "wdMergeSubTypeAccess": WdMergeSubTypeFromString = wdMergeSubTypeAccess
        Case "wdMergeSubTypeOAL": WdMergeSubTypeFromString = wdMergeSubTypeOAL
        Case "wdMergeSubTypeOLEDBWord": WdMergeSubTypeFromString = wdMergeSubTypeOLEDBWord
        Case "wdMergeSubTypeWorks": WdMergeSubTypeFromString = wdMergeSubTypeWorks
        Case "wdMergeSubTypeOLEDBText": WdMergeSubTypeFromString = wdMergeSubTypeOLEDBText
        Case "wdMergeSubTypeOutlook": WdMergeSubTypeFromString = wdMergeSubTypeOutlook
        Case "wdMergeSubTypeWord": WdMergeSubTypeFromString = wdMergeSubTypeWord
        Case "wdMergeSubTypeWord2000": WdMergeSubTypeFromString = wdMergeSubTypeWord2000
    End Select
End Function

Function WdMergeSubTypeToString(value As WdMergeSubType) As String
    Select Case value
        Case wdMergeSubTypeOther: WdMergeSubTypeToString = "wdMergeSubTypeOther"
        Case wdMergeSubTypeAccess: WdMergeSubTypeToString = "wdMergeSubTypeAccess"
        Case wdMergeSubTypeOAL: WdMergeSubTypeToString = "wdMergeSubTypeOAL"
        Case wdMergeSubTypeOLEDBWord: WdMergeSubTypeToString = "wdMergeSubTypeOLEDBWord"
        Case wdMergeSubTypeWorks: WdMergeSubTypeToString = "wdMergeSubTypeWorks"
        Case wdMergeSubTypeOLEDBText: WdMergeSubTypeToString = "wdMergeSubTypeOLEDBText"
        Case wdMergeSubTypeOutlook: WdMergeSubTypeToString = "wdMergeSubTypeOutlook"
        Case wdMergeSubTypeWord: WdMergeSubTypeToString = "wdMergeSubTypeWord"
        Case wdMergeSubTypeWord2000: WdMergeSubTypeToString = "wdMergeSubTypeWord2000"
    End Select
End Function
