Attribute VB_Name = "wXlFormatConditionOperator"
Function XlFormatConditionOperatorFromString(value As String) As XlFormatConditionOperator
    If IsNumeric(value) Then
        XlFormatConditionOperatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlBetween": XlFormatConditionOperatorFromString = xlBetween
        Case "xlNotBetween": XlFormatConditionOperatorFromString = xlNotBetween
        Case "xlEqual": XlFormatConditionOperatorFromString = xlEqual
        Case "xlNotEqual": XlFormatConditionOperatorFromString = xlNotEqual
        Case "xlGreater": XlFormatConditionOperatorFromString = xlGreater
        Case "xlLess": XlFormatConditionOperatorFromString = xlLess
        Case "xlGreaterEqual": XlFormatConditionOperatorFromString = xlGreaterEqual
        Case "xlLessEqual": XlFormatConditionOperatorFromString = xlLessEqual
    End Select
End Function

Function XlFormatConditionOperatorToString(value As XlFormatConditionOperator) As String
    Select Case value
        Case xlBetween: XlFormatConditionOperatorToString = "xlBetween"
        Case xlNotBetween: XlFormatConditionOperatorToString = "xlNotBetween"
        Case xlEqual: XlFormatConditionOperatorToString = "xlEqual"
        Case xlNotEqual: XlFormatConditionOperatorToString = "xlNotEqual"
        Case xlGreater: XlFormatConditionOperatorToString = "xlGreater"
        Case xlLess: XlFormatConditionOperatorToString = "xlLess"
        Case xlGreaterEqual: XlFormatConditionOperatorToString = "xlGreaterEqual"
        Case xlLessEqual: XlFormatConditionOperatorToString = "xlLessEqual"
    End Select
End Function
