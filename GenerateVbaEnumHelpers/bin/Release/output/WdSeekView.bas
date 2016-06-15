Attribute VB_Name = "wWdSeekView"
Function WdSeekViewFromString(value As String) As WdSeekView
    If IsNumeric(value) Then
        WdSeekViewFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSeekMainDocument": WdSeekViewFromString = wdSeekMainDocument
        Case "wdSeekPrimaryHeader": WdSeekViewFromString = wdSeekPrimaryHeader
        Case "wdSeekFirstPageHeader": WdSeekViewFromString = wdSeekFirstPageHeader
        Case "wdSeekEvenPagesHeader": WdSeekViewFromString = wdSeekEvenPagesHeader
        Case "wdSeekPrimaryFooter": WdSeekViewFromString = wdSeekPrimaryFooter
        Case "wdSeekFirstPageFooter": WdSeekViewFromString = wdSeekFirstPageFooter
        Case "wdSeekEvenPagesFooter": WdSeekViewFromString = wdSeekEvenPagesFooter
        Case "wdSeekFootnotes": WdSeekViewFromString = wdSeekFootnotes
        Case "wdSeekEndnotes": WdSeekViewFromString = wdSeekEndnotes
        Case "wdSeekCurrentPageHeader": WdSeekViewFromString = wdSeekCurrentPageHeader
        Case "wdSeekCurrentPageFooter": WdSeekViewFromString = wdSeekCurrentPageFooter
    End Select
End Function

Function WdSeekViewToString(value As WdSeekView) As String
    Select Case value
        Case wdSeekMainDocument: WdSeekViewToString = "wdSeekMainDocument"
        Case wdSeekPrimaryHeader: WdSeekViewToString = "wdSeekPrimaryHeader"
        Case wdSeekFirstPageHeader: WdSeekViewToString = "wdSeekFirstPageHeader"
        Case wdSeekEvenPagesHeader: WdSeekViewToString = "wdSeekEvenPagesHeader"
        Case wdSeekPrimaryFooter: WdSeekViewToString = "wdSeekPrimaryFooter"
        Case wdSeekFirstPageFooter: WdSeekViewToString = "wdSeekFirstPageFooter"
        Case wdSeekEvenPagesFooter: WdSeekViewToString = "wdSeekEvenPagesFooter"
        Case wdSeekFootnotes: WdSeekViewToString = "wdSeekFootnotes"
        Case wdSeekEndnotes: WdSeekViewToString = "wdSeekEndnotes"
        Case wdSeekCurrentPageHeader: WdSeekViewToString = "wdSeekCurrentPageHeader"
        Case wdSeekCurrentPageFooter: WdSeekViewToString = "wdSeekCurrentPageFooter"
    End Select
End Function
