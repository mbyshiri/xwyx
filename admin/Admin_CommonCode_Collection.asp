<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'==================================================
'��������GetArray
'��  �ã���ȡ���ӵ�ַ,��$Array$�ָ�
'��  ����ConStr ------��ȡ��ַ��ԭ�ַ�
'��  ����StartStr ------��ʼ�ַ���
'��  ����OverStr ------�����ַ���
'��  ����IncluL ------�Ƿ����StartStr
'��  ����IncluR ------�Ƿ����OverStr
'==================================================
Function GetArray(ByVal ConStr, StartStr, OverStr, IncluL, IncluR)
    If ConStr = "$False$" Or ConStr = "" Or IsNull(ConStr) = True Or StartStr = "" Or OverStr = "" Or IsNull(StartStr) = True Or IsNull(OverStr) = True Then
        GetArray = "$False$"
        Exit Function
    End If
    Dim tempStr, TempStr2
    Dim StartStr2, OverStr2
    StartStr2 = ReplaceBadChar2(StartStr)
    OverStr2 = ReplaceBadChar2(OverStr)
    regEx.Pattern = "(" & StartStr2 & ").+?(" & OverStr2 & ")"
    Set Matches = regEx.Execute(ConStr)
    For Each Match In Matches
        If tempStr <> "" Then
            tempStr = tempStr & "$Array$" & Match.value
        Else
            tempStr = Match.value
        End If
    Next
    If IncluL = False Then
        tempStr = Replace(tempStr, StartStr, "")
    End If
    If IncluR = False Then
        tempStr = Replace(tempStr, OverStr, "")
    End If
    tempStr = Replace(tempStr, """", "")
    tempStr = Replace(tempStr, "'", "")
    If tempStr = "" Then
        GetArray = "$False$"
    Else
        GetArray = tempStr
    End If
End Function
'==================================================
'��������GetPaing
'��  �ã���ȡ��ҳ
'��  ����ConStr   ------Ҫ�ҵ�����
'��  ����StartStr ------������ַͷ��
'��  ����OverStr  ------������ַβ��
'��  ����IncluL   ------�Ƿ�ͳ����ַͷ��
'��  ����IncluR   ------�Ƿ�ͳ����ַβ��
'==================================================
Function GetPaing(ByVal ConStr, StartStr, OverStr, IncluL, IncluR)
    If ConStr = "$False$" Or ConStr = "" Or StartStr = "" Or OverStr = "" Or IsNull(ConStr) = True Or IsNull(StartStr) = True Or IsNull(OverStr) = True Then
        GetPaing = "$False$"
        Exit Function
    End If
    Dim Start, Over, ConTemp, tempStr
    tempStr = LCase(ConStr)
    StartStr = LCase(StartStr)
    OverStr = LCase(OverStr)
    Over = InStr(1, tempStr, OverStr)
    If Over <= 0 Then
        GetPaing = "$False$"
        Exit Function
    Else
        If IncluR = True Then
            Over = Over + Len(OverStr)
        End If
    End If
    tempStr = Mid(tempStr, 1, Over)
    Start = InStrRev(tempStr, StartStr)
    If IncluL = False Then
        Start = Start + Len(StartStr)
    End If
    
    If Start <= 0 Or Start >= Over Then
        GetPaing = "$False$"
        Exit Function
    End If
    ConTemp = Mid(ConStr, Start, Over - Start)
    ConTemp = Trim(ConTemp)
    ConTemp = Replace(ConTemp, " ", "%20")
    ConTemp = Replace(ConTemp, ",", "")
    ConTemp = Replace(ConTemp, "'", "")
    ConTemp = Replace(ConTemp, """", "")
    ConTemp = Replace(ConTemp, ">", "")
    ConTemp = Replace(ConTemp, "<", "")
    ConTemp = Replace(ConTemp, "&nbsp;", "")
    GetPaing = ConTemp
End Function


'**************************************************
'��������CheckUrl
'��  �ã�����Ƿ��Ǿ���·������ҳ
'��  ����strUrl---Ҫ������ҳ·��
'����ֵ��True or False
'**************************************************
Function CheckUrl(ByVal strUrl)
   regEx.Pattern = "http://([\w-]+\.)+[\w-]+(/[\w-./?%&=]*)?"
   If regEx.Test(strUrl) = True Then
      CheckUrl = strUrl
   Else
      CheckUrl = False
   End If
End Function

'==================================================
'��������FpHtmlEnCode
'��  �ã��������
'��  ����fString ------�ַ���
'==================================================
Function FpHtmlEnCode(fString)
    If IsNull(fString) = False Or fString <> "" Or fString <> "$False$" Then
        fString = nohtml(fString)
        fString = FilterJS(fString)
        fString = Replace(fString, "&nbsp;", " ")
        fString = Replace(fString, "&quot;", "")
        fString = Replace(fString, "&#39;", "")
        fString = Replace(fString, ">", "")
        fString = Replace(fString, "<", "")
        fString = Replace(fString, Chr(9), " ") '&nbsp;
        fString = Replace(fString, Chr(10), "")
        fString = Replace(fString, Chr(13), "")
        fString = Replace(fString, Chr(34), "")
        fString = Replace(fString, Chr(32), " ") 'space
        fString = Replace(fString, Chr(39), "")
        fString = Replace(fString, Chr(10) & Chr(10), "")
        fString = Replace(fString, Chr(10) & Chr(13), "")
        fString = Trim(fString)
        FpHtmlEnCode = fString
    Else
        FpHtmlEnCode = "$False$"
    End If
End Function
%>
