<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Class RSA

Public PrivateKey
Public PublicKey
Public Modulus

Public Function Decode(ByVal pStrMessage)
    Dim lBytAscii
    Dim lLngIndex
    Dim lLngMaxIndex
    Dim lLngEncryptedData
    Decode = ""
    lLngMaxIndex = Len(pStrMessage)
    For lLngIndex = 1 To lLngMaxIndex Step 4
        lLngEncryptedData = HexToNumber(Mid(pStrMessage, lLngIndex, 4))
        lBytAscii = Crypt(lLngEncryptedData, PrivateKey)
        Decode = Decode & Chr(lBytAscii)
    Next
End Function

Private Function HexToNumber(ByRef pStrHex)
    HexToNumber = PE_CLng("&h" & pStrHex)
End Function

Public Function Crypt(pLngMessage, pLngKey)
    On Error Resume Next
    Dim lLngMod
    Dim lLngResult
    Dim lLngIndex
    If pLngKey Mod 2 = 0 Then
        lLngResult = 1
        For lLngIndex = 1 To pLngKey / 2
            lLngMod = (pLngMessage ^ 2) Mod Modulus
            ' Mod may error on key generation
            lLngResult = (lLngMod * lLngResult) Mod Modulus
            If Err Then Exit Function
        Next
    Else
        lLngResult = pLngMessage
        For lLngIndex = 1 To pLngKey / 2
            lLngMod = (pLngMessage ^ 2) Mod Modulus
            On Error Resume Next
            ' Mod may error on key generation
            lLngResult = (lLngMod * lLngResult) Mod Modulus
            If Err Then Exit Function
        Next
    End If
    Crypt = lLngResult
End Function

End Class
%>
