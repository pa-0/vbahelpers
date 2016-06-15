Attribute VB_Name = "wWdTextureIndex"
Function WdTextureIndexFromString(value As String) As WdTextureIndex
    If IsNumeric(value) Then
        WdTextureIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTextureNone": WdTextureIndexFromString = wdTextureNone
        Case "wdTexture2Pt5Percent": WdTextureIndexFromString = wdTexture2Pt5Percent
        Case "wdTexture5Percent": WdTextureIndexFromString = wdTexture5Percent
        Case "wdTexture7Pt5Percent": WdTextureIndexFromString = wdTexture7Pt5Percent
        Case "wdTexture10Percent": WdTextureIndexFromString = wdTexture10Percent
        Case "wdTexture12Pt5Percent": WdTextureIndexFromString = wdTexture12Pt5Percent
        Case "wdTexture15Percent": WdTextureIndexFromString = wdTexture15Percent
        Case "wdTexture17Pt5Percent": WdTextureIndexFromString = wdTexture17Pt5Percent
        Case "wdTexture20Percent": WdTextureIndexFromString = wdTexture20Percent
        Case "wdTexture22Pt5Percent": WdTextureIndexFromString = wdTexture22Pt5Percent
        Case "wdTexture25Percent": WdTextureIndexFromString = wdTexture25Percent
        Case "wdTexture27Pt5Percent": WdTextureIndexFromString = wdTexture27Pt5Percent
        Case "wdTexture30Percent": WdTextureIndexFromString = wdTexture30Percent
        Case "wdTexture32Pt5Percent": WdTextureIndexFromString = wdTexture32Pt5Percent
        Case "wdTexture35Percent": WdTextureIndexFromString = wdTexture35Percent
        Case "wdTexture37Pt5Percent": WdTextureIndexFromString = wdTexture37Pt5Percent
        Case "wdTexture40Percent": WdTextureIndexFromString = wdTexture40Percent
        Case "wdTexture42Pt5Percent": WdTextureIndexFromString = wdTexture42Pt5Percent
        Case "wdTexture45Percent": WdTextureIndexFromString = wdTexture45Percent
        Case "wdTexture47Pt5Percent": WdTextureIndexFromString = wdTexture47Pt5Percent
        Case "wdTexture50Percent": WdTextureIndexFromString = wdTexture50Percent
        Case "wdTexture52Pt5Percent": WdTextureIndexFromString = wdTexture52Pt5Percent
        Case "wdTexture55Percent": WdTextureIndexFromString = wdTexture55Percent
        Case "wdTexture57Pt5Percent": WdTextureIndexFromString = wdTexture57Pt5Percent
        Case "wdTexture60Percent": WdTextureIndexFromString = wdTexture60Percent
        Case "wdTexture62Pt5Percent": WdTextureIndexFromString = wdTexture62Pt5Percent
        Case "wdTexture65Percent": WdTextureIndexFromString = wdTexture65Percent
        Case "wdTexture67Pt5Percent": WdTextureIndexFromString = wdTexture67Pt5Percent
        Case "wdTexture70Percent": WdTextureIndexFromString = wdTexture70Percent
        Case "wdTexture72Pt5Percent": WdTextureIndexFromString = wdTexture72Pt5Percent
        Case "wdTexture75Percent": WdTextureIndexFromString = wdTexture75Percent
        Case "wdTexture77Pt5Percent": WdTextureIndexFromString = wdTexture77Pt5Percent
        Case "wdTexture80Percent": WdTextureIndexFromString = wdTexture80Percent
        Case "wdTexture82Pt5Percent": WdTextureIndexFromString = wdTexture82Pt5Percent
        Case "wdTexture85Percent": WdTextureIndexFromString = wdTexture85Percent
        Case "wdTexture87Pt5Percent": WdTextureIndexFromString = wdTexture87Pt5Percent
        Case "wdTexture90Percent": WdTextureIndexFromString = wdTexture90Percent
        Case "wdTexture92Pt5Percent": WdTextureIndexFromString = wdTexture92Pt5Percent
        Case "wdTexture95Percent": WdTextureIndexFromString = wdTexture95Percent
        Case "wdTexture97Pt5Percent": WdTextureIndexFromString = wdTexture97Pt5Percent
        Case "wdTextureSolid": WdTextureIndexFromString = wdTextureSolid
        Case "wdTextureDiagonalCross": WdTextureIndexFromString = wdTextureDiagonalCross
        Case "wdTextureCross": WdTextureIndexFromString = wdTextureCross
        Case "wdTextureDiagonalUp": WdTextureIndexFromString = wdTextureDiagonalUp
        Case "wdTextureDiagonalDown": WdTextureIndexFromString = wdTextureDiagonalDown
        Case "wdTextureVertical": WdTextureIndexFromString = wdTextureVertical
        Case "wdTextureHorizontal": WdTextureIndexFromString = wdTextureHorizontal
        Case "wdTextureDarkDiagonalCross": WdTextureIndexFromString = wdTextureDarkDiagonalCross
        Case "wdTextureDarkCross": WdTextureIndexFromString = wdTextureDarkCross
        Case "wdTextureDarkDiagonalUp": WdTextureIndexFromString = wdTextureDarkDiagonalUp
        Case "wdTextureDarkDiagonalDown": WdTextureIndexFromString = wdTextureDarkDiagonalDown
        Case "wdTextureDarkVertical": WdTextureIndexFromString = wdTextureDarkVertical
        Case "wdTextureDarkHorizontal": WdTextureIndexFromString = wdTextureDarkHorizontal
    End Select
End Function

Function WdTextureIndexToString(value As WdTextureIndex) As String
    Select Case value
        Case wdTextureNone: WdTextureIndexToString = "wdTextureNone"
        Case wdTexture2Pt5Percent: WdTextureIndexToString = "wdTexture2Pt5Percent"
        Case wdTexture5Percent: WdTextureIndexToString = "wdTexture5Percent"
        Case wdTexture7Pt5Percent: WdTextureIndexToString = "wdTexture7Pt5Percent"
        Case wdTexture10Percent: WdTextureIndexToString = "wdTexture10Percent"
        Case wdTexture12Pt5Percent: WdTextureIndexToString = "wdTexture12Pt5Percent"
        Case wdTexture15Percent: WdTextureIndexToString = "wdTexture15Percent"
        Case wdTexture17Pt5Percent: WdTextureIndexToString = "wdTexture17Pt5Percent"
        Case wdTexture20Percent: WdTextureIndexToString = "wdTexture20Percent"
        Case wdTexture22Pt5Percent: WdTextureIndexToString = "wdTexture22Pt5Percent"
        Case wdTexture25Percent: WdTextureIndexToString = "wdTexture25Percent"
        Case wdTexture27Pt5Percent: WdTextureIndexToString = "wdTexture27Pt5Percent"
        Case wdTexture30Percent: WdTextureIndexToString = "wdTexture30Percent"
        Case wdTexture32Pt5Percent: WdTextureIndexToString = "wdTexture32Pt5Percent"
        Case wdTexture35Percent: WdTextureIndexToString = "wdTexture35Percent"
        Case wdTexture37Pt5Percent: WdTextureIndexToString = "wdTexture37Pt5Percent"
        Case wdTexture40Percent: WdTextureIndexToString = "wdTexture40Percent"
        Case wdTexture42Pt5Percent: WdTextureIndexToString = "wdTexture42Pt5Percent"
        Case wdTexture45Percent: WdTextureIndexToString = "wdTexture45Percent"
        Case wdTexture47Pt5Percent: WdTextureIndexToString = "wdTexture47Pt5Percent"
        Case wdTexture50Percent: WdTextureIndexToString = "wdTexture50Percent"
        Case wdTexture52Pt5Percent: WdTextureIndexToString = "wdTexture52Pt5Percent"
        Case wdTexture55Percent: WdTextureIndexToString = "wdTexture55Percent"
        Case wdTexture57Pt5Percent: WdTextureIndexToString = "wdTexture57Pt5Percent"
        Case wdTexture60Percent: WdTextureIndexToString = "wdTexture60Percent"
        Case wdTexture62Pt5Percent: WdTextureIndexToString = "wdTexture62Pt5Percent"
        Case wdTexture65Percent: WdTextureIndexToString = "wdTexture65Percent"
        Case wdTexture67Pt5Percent: WdTextureIndexToString = "wdTexture67Pt5Percent"
        Case wdTexture70Percent: WdTextureIndexToString = "wdTexture70Percent"
        Case wdTexture72Pt5Percent: WdTextureIndexToString = "wdTexture72Pt5Percent"
        Case wdTexture75Percent: WdTextureIndexToString = "wdTexture75Percent"
        Case wdTexture77Pt5Percent: WdTextureIndexToString = "wdTexture77Pt5Percent"
        Case wdTexture80Percent: WdTextureIndexToString = "wdTexture80Percent"
        Case wdTexture82Pt5Percent: WdTextureIndexToString = "wdTexture82Pt5Percent"
        Case wdTexture85Percent: WdTextureIndexToString = "wdTexture85Percent"
        Case wdTexture87Pt5Percent: WdTextureIndexToString = "wdTexture87Pt5Percent"
        Case wdTexture90Percent: WdTextureIndexToString = "wdTexture90Percent"
        Case wdTexture92Pt5Percent: WdTextureIndexToString = "wdTexture92Pt5Percent"
        Case wdTexture95Percent: WdTextureIndexToString = "wdTexture95Percent"
        Case wdTexture97Pt5Percent: WdTextureIndexToString = "wdTexture97Pt5Percent"
        Case wdTextureSolid: WdTextureIndexToString = "wdTextureSolid"
        Case wdTextureDiagonalCross: WdTextureIndexToString = "wdTextureDiagonalCross"
        Case wdTextureCross: WdTextureIndexToString = "wdTextureCross"
        Case wdTextureDiagonalUp: WdTextureIndexToString = "wdTextureDiagonalUp"
        Case wdTextureDiagonalDown: WdTextureIndexToString = "wdTextureDiagonalDown"
        Case wdTextureVertical: WdTextureIndexToString = "wdTextureVertical"
        Case wdTextureHorizontal: WdTextureIndexToString = "wdTextureHorizontal"
        Case wdTextureDarkDiagonalCross: WdTextureIndexToString = "wdTextureDarkDiagonalCross"
        Case wdTextureDarkCross: WdTextureIndexToString = "wdTextureDarkCross"
        Case wdTextureDarkDiagonalUp: WdTextureIndexToString = "wdTextureDarkDiagonalUp"
        Case wdTextureDarkDiagonalDown: WdTextureIndexToString = "wdTextureDarkDiagonalDown"
        Case wdTextureDarkVertical: WdTextureIndexToString = "wdTextureDarkVertical"
        Case wdTextureDarkHorizontal: WdTextureIndexToString = "wdTextureDarkHorizontal"
    End Select
End Function
