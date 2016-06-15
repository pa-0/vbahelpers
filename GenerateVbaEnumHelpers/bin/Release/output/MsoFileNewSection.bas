Attribute VB_Name = "wMsoFileNewSection"
Function MsoFileNewSectionFromString(value As String) As MsoFileNewSection
    If IsNumeric(value) Then
        MsoFileNewSectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoOpenDocument": MsoFileNewSectionFromString = msoOpenDocument
        Case "msoNew": MsoFileNewSectionFromString = msoNew
        Case "msoNewfromExistingFile": MsoFileNewSectionFromString = msoNewfromExistingFile
        Case "msoNewfromTemplate": MsoFileNewSectionFromString = msoNewfromTemplate
        Case "msoBottomSection": MsoFileNewSectionFromString = msoBottomSection
    End Select
End Function

Function MsoFileNewSectionToString(value As MsoFileNewSection) As String
    Select Case value
        Case msoOpenDocument: MsoFileNewSectionToString = "msoOpenDocument"
        Case msoNew: MsoFileNewSectionToString = "msoNew"
        Case msoNewfromExistingFile: MsoFileNewSectionToString = "msoNewfromExistingFile"
        Case msoNewfromTemplate: MsoFileNewSectionToString = "msoNewfromTemplate"
        Case msoBottomSection: MsoFileNewSectionToString = "msoBottomSection"
    End Select
End Function
