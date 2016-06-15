Attribute VB_Name = "wWdBrowseTarget"
Function WdBrowseTargetFromString(value As String) As WdBrowseTarget
    If IsNumeric(value) Then
        WdBrowseTargetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBrowsePage": WdBrowseTargetFromString = wdBrowsePage
        Case "wdBrowseSection": WdBrowseTargetFromString = wdBrowseSection
        Case "wdBrowseComment": WdBrowseTargetFromString = wdBrowseComment
        Case "wdBrowseFootnote": WdBrowseTargetFromString = wdBrowseFootnote
        Case "wdBrowseEndnote": WdBrowseTargetFromString = wdBrowseEndnote
        Case "wdBrowseField": WdBrowseTargetFromString = wdBrowseField
        Case "wdBrowseTable": WdBrowseTargetFromString = wdBrowseTable
        Case "wdBrowseGraphic": WdBrowseTargetFromString = wdBrowseGraphic
        Case "wdBrowseHeading": WdBrowseTargetFromString = wdBrowseHeading
        Case "wdBrowseEdit": WdBrowseTargetFromString = wdBrowseEdit
        Case "wdBrowseFind": WdBrowseTargetFromString = wdBrowseFind
        Case "wdBrowseGoTo": WdBrowseTargetFromString = wdBrowseGoTo
    End Select
End Function

Function WdBrowseTargetToString(value As WdBrowseTarget) As String
    Select Case value
        Case wdBrowsePage: WdBrowseTargetToString = "wdBrowsePage"
        Case wdBrowseSection: WdBrowseTargetToString = "wdBrowseSection"
        Case wdBrowseComment: WdBrowseTargetToString = "wdBrowseComment"
        Case wdBrowseFootnote: WdBrowseTargetToString = "wdBrowseFootnote"
        Case wdBrowseEndnote: WdBrowseTargetToString = "wdBrowseEndnote"
        Case wdBrowseField: WdBrowseTargetToString = "wdBrowseField"
        Case wdBrowseTable: WdBrowseTargetToString = "wdBrowseTable"
        Case wdBrowseGraphic: WdBrowseTargetToString = "wdBrowseGraphic"
        Case wdBrowseHeading: WdBrowseTargetToString = "wdBrowseHeading"
        Case wdBrowseEdit: WdBrowseTargetToString = "wdBrowseEdit"
        Case wdBrowseFind: WdBrowseTargetToString = "wdBrowseFind"
        Case wdBrowseGoTo: WdBrowseTargetToString = "wdBrowseGoTo"
    End Select
End Function
