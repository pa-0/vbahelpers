Attribute VB_Name = "wWdBuiltInProperty"
Function WdBuiltInPropertyFromString(value As String) As WdBuiltInProperty
    If IsNumeric(value) Then
        WdBuiltInPropertyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPropertyTitle": WdBuiltInPropertyFromString = wdPropertyTitle
        Case "wdPropertySubject": WdBuiltInPropertyFromString = wdPropertySubject
        Case "wdPropertyAuthor": WdBuiltInPropertyFromString = wdPropertyAuthor
        Case "wdPropertyKeywords": WdBuiltInPropertyFromString = wdPropertyKeywords
        Case "wdPropertyComments": WdBuiltInPropertyFromString = wdPropertyComments
        Case "wdPropertyTemplate": WdBuiltInPropertyFromString = wdPropertyTemplate
        Case "wdPropertyLastAuthor": WdBuiltInPropertyFromString = wdPropertyLastAuthor
        Case "wdPropertyRevision": WdBuiltInPropertyFromString = wdPropertyRevision
        Case "wdPropertyAppName": WdBuiltInPropertyFromString = wdPropertyAppName
        Case "wdPropertyTimeLastPrinted": WdBuiltInPropertyFromString = wdPropertyTimeLastPrinted
        Case "wdPropertyTimeCreated": WdBuiltInPropertyFromString = wdPropertyTimeCreated
        Case "wdPropertyTimeLastSaved": WdBuiltInPropertyFromString = wdPropertyTimeLastSaved
        Case "wdPropertyVBATotalEdit": WdBuiltInPropertyFromString = wdPropertyVBATotalEdit
        Case "wdPropertyPages": WdBuiltInPropertyFromString = wdPropertyPages
        Case "wdPropertyWords": WdBuiltInPropertyFromString = wdPropertyWords
        Case "wdPropertyCharacters": WdBuiltInPropertyFromString = wdPropertyCharacters
        Case "wdPropertySecurity": WdBuiltInPropertyFromString = wdPropertySecurity
        Case "wdPropertyCategory": WdBuiltInPropertyFromString = wdPropertyCategory
        Case "wdPropertyFormat": WdBuiltInPropertyFromString = wdPropertyFormat
        Case "wdPropertyManager": WdBuiltInPropertyFromString = wdPropertyManager
        Case "wdPropertyCompany": WdBuiltInPropertyFromString = wdPropertyCompany
        Case "wdPropertyBytes": WdBuiltInPropertyFromString = wdPropertyBytes
        Case "wdPropertyLines": WdBuiltInPropertyFromString = wdPropertyLines
        Case "wdPropertyParas": WdBuiltInPropertyFromString = wdPropertyParas
        Case "wdPropertySlides": WdBuiltInPropertyFromString = wdPropertySlides
        Case "wdPropertyNotes": WdBuiltInPropertyFromString = wdPropertyNotes
        Case "wdPropertyHiddenSlides": WdBuiltInPropertyFromString = wdPropertyHiddenSlides
        Case "wdPropertyMMClips": WdBuiltInPropertyFromString = wdPropertyMMClips
        Case "wdPropertyHyperlinkBase": WdBuiltInPropertyFromString = wdPropertyHyperlinkBase
        Case "wdPropertyCharsWSpaces": WdBuiltInPropertyFromString = wdPropertyCharsWSpaces
    End Select
End Function

Function WdBuiltInPropertyToString(value As WdBuiltInProperty) As String
    Select Case value
        Case wdPropertyTitle: WdBuiltInPropertyToString = "wdPropertyTitle"
        Case wdPropertySubject: WdBuiltInPropertyToString = "wdPropertySubject"
        Case wdPropertyAuthor: WdBuiltInPropertyToString = "wdPropertyAuthor"
        Case wdPropertyKeywords: WdBuiltInPropertyToString = "wdPropertyKeywords"
        Case wdPropertyComments: WdBuiltInPropertyToString = "wdPropertyComments"
        Case wdPropertyTemplate: WdBuiltInPropertyToString = "wdPropertyTemplate"
        Case wdPropertyLastAuthor: WdBuiltInPropertyToString = "wdPropertyLastAuthor"
        Case wdPropertyRevision: WdBuiltInPropertyToString = "wdPropertyRevision"
        Case wdPropertyAppName: WdBuiltInPropertyToString = "wdPropertyAppName"
        Case wdPropertyTimeLastPrinted: WdBuiltInPropertyToString = "wdPropertyTimeLastPrinted"
        Case wdPropertyTimeCreated: WdBuiltInPropertyToString = "wdPropertyTimeCreated"
        Case wdPropertyTimeLastSaved: WdBuiltInPropertyToString = "wdPropertyTimeLastSaved"
        Case wdPropertyVBATotalEdit: WdBuiltInPropertyToString = "wdPropertyVBATotalEdit"
        Case wdPropertyPages: WdBuiltInPropertyToString = "wdPropertyPages"
        Case wdPropertyWords: WdBuiltInPropertyToString = "wdPropertyWords"
        Case wdPropertyCharacters: WdBuiltInPropertyToString = "wdPropertyCharacters"
        Case wdPropertySecurity: WdBuiltInPropertyToString = "wdPropertySecurity"
        Case wdPropertyCategory: WdBuiltInPropertyToString = "wdPropertyCategory"
        Case wdPropertyFormat: WdBuiltInPropertyToString = "wdPropertyFormat"
        Case wdPropertyManager: WdBuiltInPropertyToString = "wdPropertyManager"
        Case wdPropertyCompany: WdBuiltInPropertyToString = "wdPropertyCompany"
        Case wdPropertyBytes: WdBuiltInPropertyToString = "wdPropertyBytes"
        Case wdPropertyLines: WdBuiltInPropertyToString = "wdPropertyLines"
        Case wdPropertyParas: WdBuiltInPropertyToString = "wdPropertyParas"
        Case wdPropertySlides: WdBuiltInPropertyToString = "wdPropertySlides"
        Case wdPropertyNotes: WdBuiltInPropertyToString = "wdPropertyNotes"
        Case wdPropertyHiddenSlides: WdBuiltInPropertyToString = "wdPropertyHiddenSlides"
        Case wdPropertyMMClips: WdBuiltInPropertyToString = "wdPropertyMMClips"
        Case wdPropertyHyperlinkBase: WdBuiltInPropertyToString = "wdPropertyHyperlinkBase"
        Case wdPropertyCharsWSpaces: WdBuiltInPropertyToString = "wdPropertyCharsWSpaces"
    End Select
End Function
