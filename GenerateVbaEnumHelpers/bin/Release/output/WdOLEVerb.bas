Attribute VB_Name = "wWdOLEVerb"
Function WdOLEVerbFromString(value As String) As WdOLEVerb
    If IsNumeric(value) Then
        WdOLEVerbFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOLEVerbPrimary": WdOLEVerbFromString = wdOLEVerbPrimary
        Case "wdOLEVerbDiscardUndoState": WdOLEVerbFromString = wdOLEVerbDiscardUndoState
        Case "wdOLEVerbInPlaceActivate": WdOLEVerbFromString = wdOLEVerbInPlaceActivate
        Case "wdOLEVerbUIActivate": WdOLEVerbFromString = wdOLEVerbUIActivate
        Case "wdOLEVerbHide": WdOLEVerbFromString = wdOLEVerbHide
        Case "wdOLEVerbOpen": WdOLEVerbFromString = wdOLEVerbOpen
        Case "wdOLEVerbShow": WdOLEVerbFromString = wdOLEVerbShow
    End Select
End Function

Function WdOLEVerbToString(value As WdOLEVerb) As String
    Select Case value
        Case wdOLEVerbPrimary: WdOLEVerbToString = "wdOLEVerbPrimary"
        Case wdOLEVerbDiscardUndoState: WdOLEVerbToString = "wdOLEVerbDiscardUndoState"
        Case wdOLEVerbInPlaceActivate: WdOLEVerbToString = "wdOLEVerbInPlaceActivate"
        Case wdOLEVerbUIActivate: WdOLEVerbToString = "wdOLEVerbUIActivate"
        Case wdOLEVerbHide: WdOLEVerbToString = "wdOLEVerbHide"
        Case wdOLEVerbOpen: WdOLEVerbToString = "wdOLEVerbOpen"
        Case wdOLEVerbShow: WdOLEVerbToString = "wdOLEVerbShow"
    End Select
End Function
