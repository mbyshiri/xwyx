<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const BASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private sBASE_64_CHARACTERS


Public Function Base64encode(ByVal asContents)
    asContents = strUnicode2Ansi(asContents)
    
    Dim lnPosition
    Dim lsResult
    Dim Char1
    Dim Char2
    Dim Char3
    Dim Char4
    Dim Byte1
    Dim Byte2
    Dim Byte3
    Dim SaveBits1
    Dim SaveBits2
    Dim lsGroupBinary
    Dim lsGroup64
    Dim M4, len1, len2
    
    len1 = LenB(asContents)
    If len1 < 1 Then
        Base64encode = ""
        Exit Function
    End If

    M4 = len1 Mod 3
    If M4 > 0 Then asContents = asContents & String(3 - M4, Chr(0))
    '补足位数是为了便于计算

    If M4 > 0 Then
        len1 = len1 + (3 - M4)
        len2 = len1 - 3
    Else
        len2 = len1
    End If

    lsResult = ""
    sBASE_64_CHARACTERS = strUnicode2Ansi(BASE_64_CHARACTERS)
    
    For lnPosition = 1 To len2 Step 3
        lsGroup64 = ""
        lsGroupBinary = MidB(asContents, lnPosition, 3)

        Byte1 = AscB(MidB(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3
        Byte2 = AscB(MidB(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15
        Byte3 = AscB(MidB(lsGroupBinary, 3, 1))

        Char1 = MidB(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
        Char2 = MidB(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
        Char3 = MidB(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
        Char4 = MidB(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)
        lsGroup64 = Char1 & Char2 & Char3 & Char4

        lsResult = lsResult & lsGroup64
    Next

    '处理最后剩余的几个字符
    If M4 > 0 Then
        lsGroup64 = ""
        lsGroupBinary = MidB(asContents, len2 + 1, 3)

        Byte1 = AscB(MidB(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3
        Byte2 = AscB(MidB(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15
        Byte3 = AscB(MidB(lsGroupBinary, 3, 1))

        Char1 = MidB(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
        Char2 = MidB(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
        Char3 = MidB(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)

        If M4 = 1 Then
            lsGroup64 = Char1 & Char2 & ChrB(61) & ChrB(61) '用=号补足位数
        Else
            lsGroup64 = Char1 & Char2 & Char3 & ChrB(61) '用=号补足位数
        End If

        lsResult = lsResult & lsGroup64
    End If

    Base64encode = strAnsi2Unicode(lsResult)

End Function


Public Function Base64decode(ByVal asContents)
    asContents = strUnicode2Ansi(asContents)
    
    Dim lsResult
    Dim lnPosition
    Dim lsGroup64, lsGroupBinary
    Dim Char1, Char2, Char3, Char4
    Dim Byte1, Byte2, Byte3
    Dim M4, len1, len2

    len1 = LenB(asContents)
    M4 = len1 Mod 4

    If len1 < 1 Or M4 > 0 Then
        '字符串长度应当是4的倍数
        Base64decode = ""
        Exit Function
    End If

    '判断最后一位是不是 = 号
    '判断倒数第二位是不是 = 号
    '这里m4表示最后剩余的需要单独处理的字符个数
    If MidB(asContents, len1, 1) = ChrB(61) Then M4 = 3
    If MidB(asContents, len1 - 1, 1) = ChrB(61) Then M4 = 2

    If M4 = 0 Then
        len2 = len1
    Else
        len2 = len1 - 4
    End If
    
    sBASE_64_CHARACTERS = strUnicode2Ansi(BASE_64_CHARACTERS)
    
    For lnPosition = 1 To len2 Step 4
        lsGroupBinary = ""
        lsGroup64 = MidB(asContents, lnPosition, 4)
        Char1 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 1, 1)) - 1
        Char2 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 2, 1)) - 1
        Char3 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 3, 1)) - 1
        Char4 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 4, 1)) - 1
        Byte1 = ChrB(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
        Byte2 = lsGroupBinary & ChrB(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
        Byte3 = ChrB((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
        lsGroupBinary = Byte1 & Byte2 & Byte3

        lsResult = lsResult & lsGroupBinary
    Next

    '处理最后剩余的几个字符
    If M4 > 0 Then
        lsGroupBinary = ""
        lsGroup64 = MidB(asContents, len2 + 1, M4) & ChrB(65) 'chr(65)=A，转换成值为0
        If M4 = 2 Then '补足4位，是为了便于计算
            lsGroup64 = lsGroup64 & ChrB(65)
        End If
        Char1 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 1, 1)) - 1
        Char2 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 2, 1)) - 1
        Char3 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 3, 1)) - 1
        Char4 = InStrB(sBASE_64_CHARACTERS, MidB(lsGroup64, 4, 1)) - 1
        Byte1 = ChrB(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
        Byte2 = lsGroupBinary & ChrB(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
        Byte3 = ChrB((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))

        If M4 = 2 Then
            lsGroupBinary = Byte1
        ElseIf M4 = 3 Then
            lsGroupBinary = Byte1 & Byte2
        End If

        lsResult = lsResult & lsGroupBinary
    End If

    Base64decode = strAnsi2Unicode(lsResult)

End Function

Private Function strUnicodeLen(ByVal asContents)
    '计算unicode字符串的Ansi编码的长度
    Dim asContents1
    Dim len1
    Dim k
    Dim i
    Dim asc1
    
    asContents1 = "a" & asContents
    len1 = Len(asContents1)
    k = 0
    For i = 1 To len1
        asc1 = Asc(Mid(asContents1, i, 1))
        If asc1 < 0 Then asc1 = 65536 + asc1
        If asc1 > 255 Then
            k = k + 2
            Else
            k = k + 1
        End If
    Next
    strUnicodeLen = k - 1
End Function

Private Function strUnicode2Ansi(ByVal asContents)
    '将Unicode编码的字符串，转换成Ansi编码的字符串
    Dim len1
    Dim i
    Dim VarChar
    Dim varAsc
    Dim varHex, varlow, varhigh
    
    strUnicode2Ansi = ""
    len1 = Len(asContents)
    For i = 1 To len1
        VarChar = Mid(asContents, i, 1)
        varAsc = Asc(VarChar)
        If varAsc < 0 Then varAsc = varAsc + 65536
        If varAsc > 255 Then
            varHex = Hex(varAsc)
            varlow = Left(varHex, 2)
            varhigh = Right(varHex, 2)
            strUnicode2Ansi = strUnicode2Ansi & ChrB("&H" & varlow) & ChrB("&H" & varhigh)
        Else
            strUnicode2Ansi = strUnicode2Ansi & ChrB(varAsc)
        End If
    Next
End Function

Private Function strAnsi2Unicode(asContents)
    '将Ansi编码的字符串，转换成Unicode编码的字符串
    Dim len1
    Dim i
    Dim VarChar
    Dim varAsc
    
    strAnsi2Unicode = ""
    len1 = LenB(asContents)
    If len1 = 0 Then Exit Function
    For i = 1 To len1
        VarChar = MidB(asContents, i, 1)
        varAsc = AscB(VarChar)
        If varAsc > 127 Then
            strAnsi2Unicode = strAnsi2Unicode & Chr(AscW(MidB(asContents, i + 1, 1) & VarChar))
            i = i + 1
        Else
            strAnsi2Unicode = strAnsi2Unicode & Chr(varAsc)
        End If
    Next
End Function
%>
