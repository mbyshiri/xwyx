<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5_New.asp"-->
<!--#include file="UpdateOrder.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************
Const IsMessageShow = False
Const PlatformID = 8  '֧����
Call CheckPlatformID(PlatformID)
Dim PaySuccess
PaySuccess = False

Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_moneytype
Dim md5string

Dim Partner
Dim ArrMD5Key
If InStr(MD5Key, "|") > 0 Then
    ArrMD5Key = Split(MD5Key, "|")
    If UBound(ArrMD5Key) = 1 Then
        MD5Key = ArrMD5Key(0)
        Partner = ArrMD5Key(1)
    End If
End If

Dim alipayNotifyURL, ResponseTxt, returnTxt
Dim PE_Md5
Set PE_Md5 = New Md5_Class
v_mid = AccountsID
If Partner = "" Then  '�ɽӿ�
    Dim myalipayEmail   '�̻���֧����Email
    Dim myAlipayObj
    Dim strMd5

    Dim msg_id, buyer_email, buyer_name, buyer_address, buyer_zipcode, buyer_tel, buyer_mobile, s_date, ac

    returnTxt = "Error"    '���ظ�֧����֪ͨ�ӿڵĽ��
    alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?" '֧������ѯ�ӿ�URL
    myalipayEmail = v_mid    '�̻���֧����Email
   
    msg_id = DelStr(Request("msg_id"))             '֪ͨ����ID
    v_oid = DelStr(Request("order_no"))            '�̻�������
    v_amount = DelStr(Request("gross"))            '���
    buyer_email = DelStr(Request("buyer_email"))   '���Email
    buyer_name = DelStr(Request("buyer_name"))           '�������
    buyer_address = DelStr(Request("buyer_address"))     '��ҵ�ַ
    buyer_zipcode = DelStr(Request("buyer_zipcode"))    '����ʱ�
    buyer_tel = DelStr(Request("buyer_tel"))            '��ҵ绰
    buyer_mobile = DelStr(Request("buyer_mobile"))       '����ֻ�
    Action = DelStr(Request("action"))                    '֪ͨ����
    s_date = DelStr(Request("date"))                       '֪ͨ��������
    ac = DelStr(Request("ac"))                           'MD5ժҪ
    
    If Not IsEmpty(msg_id) Then strMd5 = "msg_id" & msg_id
    If Not IsEmpty(v_oid) Then strMd5 = strMd5 & "order_no" & v_oid
    If Not IsEmpty(v_amount) Then strMd5 = strMd5 & "gross" & v_amount
    If Not IsEmpty(buyer_email) Then strMd5 = strMd5 & "buyer_email" & buyer_email
    If Not IsEmpty(buyer_name) Then strMd5 = strMd5 & "buyer_name" & buyer_name
    If Not IsEmpty(buyer_address) Then strMd5 = strMd5 & "buyer_address" & buyer_address
    If Not IsEmpty(buyer_zipcode) Then strMd5 = strMd5 & "buyer_zipcode" & buyer_zipcode
    If Not IsEmpty(buyer_tel) Then strMd5 = strMd5 & "buyer_tel" & buyer_tel
    If Not IsEmpty(buyer_mobile) Then strMd5 = strMd5 & "buyer_mobile" & buyer_mobile
    If Not IsEmpty(Action) Then strMd5 = strMd5 & "action" & Action
    If Not IsEmpty(s_date) Then strMd5 = strMd5 & "date" & s_date
    strMd5 = strMd5 & MD5Key

    '�����Ƿ�֧����������֪ͨ
     alipayNotifyURL = alipayNotifyURL & "msg_id=" & msg_id & "&email=" & myalipayEmail & "&order_no=" & v_oid
     ResponseTxt = GetBody(alipayNotifyURL) '����֧������ѯ�ӿڣ����֧������������true��false����������ж��Ƿ���֧����������������

    Select Case Action
    Case "test"   '�����̻���վURL�Ƿ���ȷ��װ
        returnTxt = "Y"
    Case "sendOff"   '����֪ͨ
        returnTxt = "N"
        If LCase(PE_Md5.MD5(strMd5)) = LCase(ac) Then
            If (ResponseTxt = "true") Or (ResponseTxt = "false") Then
                returnTxt = "Y"
                '�����������̻�ϵͳ��Ķ������ݣ�����Ѿ���������returnTxt��ΪY������ΪN
                Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 2, True, False)
            Else
                '�Ƿ����ݣ���������
            End If
        End If
    Case "checkOut"   '���׽���֪ͨ
        returnTxt = "N"
        If LCase(PE_Md5.MD5(strMd5)) = LCase(ac) Then
            If (ResponseTxt = "true") Or (ResponseTxt = "false") Then
                '�����������̻�ϵͳ��Ķ������ݣ�������ݸ��³ɹ�����returnTxt��ΪY������ΪN
                returnTxt = "Y"
                PaySuccess = True
            Else
                '�Ƿ����ݣ���������
                returnTxt = "N"
            End If
        End If
    End Select
    Response.Write returnTxt
Else   '�½ӿ�
    Dim trade_status, sign, MySign, Retrieval
    Dim mystr, Count, i, minmax, minmaxSlot, j, mark, temp, value, md5str, notify_id
    
    v_oid = DelStr(Request("out_trade_no"))            '�̻�������
    trade_status = DelStr(Request("trade_status"))
    sign = DelStr(Request("sign"))
    v_amount = DelStr(Request("price"))
    notify_id = Request.Form("notify_id")


    alipayNotifyURL = "https://www.alipay.com/cooperate/gateway.do?"

    alipayNotifyURL = alipayNotifyURL & "service=notify_verify&partner=" & Partner & "&notify_id=" & notify_id
    Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
    Retrieval.Open "GET", alipayNotifyURL, False, "", ""
    Retrieval.Send
    ResponseTxt = Retrieval.ResponseText
    Set Retrieval = Nothing

                
    '��ȡPOST�����Ĳ���
    mystr = Split(URLDecode(Request.Form), "&")
    Count = UBound(mystr)

    '�Բ�������
    For i = Count To 0 Step -1
        minmax = mystr(0)
        minmaxSlot = 0
        For j = 1 To i
            mark = (mystr(j) > minmax)
            If mark Then
                minmax = mystr(j)
                minmaxSlot = j
            End If
        Next

        If minmaxSlot <> i Then
            temp = mystr(minmaxSlot)
            mystr(minmaxSlot) = mystr(i)
            mystr(i) = temp
        End If
    Next

    '����md5ժҪ�ַ���
    For j = 0 To Count Step 1
        value = Split(mystr(j), "=")
        If value(1) <> "" And value(0) <> "sign" And value(0) <> "sign_type" Then
            If j = Count Then
                md5str = md5str & mystr(j)
            Else
                md5str = md5str & mystr(j) & "&"
            End If
        End If
    Next

    md5str = md5str & MD5Key
    '����md5ժҪ
    MySign = PE_Md5.MD5(md5str)

    '�ȴ���Ҹ���
    Select Case trade_status
    Case "WAIT_BUYER_PAY"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
        Else
            returnTxt = "fail"
        End If

    '��Ҹ���ɹ�,�ȴ����ҷ���
    Case "WAIT_SELLER_SEND_GOODS"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
            Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 2, False, False)  'ֻ����֧����¼״̬�������·���״̬�Ͷ���״̬
        Else
            returnTxt = "fail"
        End If

    '�ȴ����ȷ���ջ�
    Case "WAIT_BUYER_CONFIRM_GOODS"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
            Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 4, True, False)  '����֧����¼״̬�ͷ���״̬�������¶���״̬
        Else
            returnTxt = "fail"
        End If

    '���׳ɹ�����
    Case "TRADE_FINISHED"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
            PaySuccess = True                '���׳ɹ������¶���
        Else
            returnTxt = "fail"
        End If

    '��������״̬֪ͨ���
    Case Else
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
        Else
            returnTxt = "fail"
        End If
    End Select
    Response.Write returnTxt
End If
If PaySuccess = True Then
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
End If

Call CloseConn

Function DelStr(str)
    If IsNull(str) Or IsEmpty(str) Then
        str = ""
    End If
    DelStr = Replace(str, ";", "")
    DelStr = Replace(DelStr, "'", "")
    DelStr = Replace(DelStr, "&", "")
    DelStr = Replace(DelStr, " ", "")
    DelStr = Replace(DelStr, "��", "")
    DelStr = Replace(DelStr, "%20", "")
    DelStr = Replace(DelStr, "--", "")
    DelStr = Replace(DelStr, "==", "")
    DelStr = Replace(DelStr, "<", "")
    DelStr = Replace(DelStr, ">", "")
    DelStr = Replace(DelStr, "%", "")
End Function

'ȡ������󷵻ص�html Stream
Function GetBody(strURL)
    On Error Resume Next
    Dim Retrieval
    Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
    Retrieval.Open "GET", strURL, False, "", ""
    Retrieval.Send
    GetBody = Retrieval.ResponseText
    Set Retrieval = Nothing
End Function

'��post���ݹ����Ĳ�����urldecode���봦��(֧�������½ӿ�)
Function URLDecode(enStr)
    Dim deStr
    Dim c, i, v
    deStr = ""
    For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
            v = eval("&h" + Mid(enStr, i + 1, 2))
            If v < 128 Then
                deStr = deStr & Chr(v)
                i = i + 2
            Else
                If isvalidhex(Mid(enStr, i, 3)) Then
                    If isvalidhex(Mid(enStr, i + 3, 3)) Then
                        v = eval("&h" + Mid(enStr, i + 1, 2) + Mid(enStr, i + 4, 2))
                        deStr = deStr & Chr(v)
                        i = i + 5
                    Else
                        v = eval("&h" + Mid(enStr, i + 1, 2) + CStr(Hex(Asc(Mid(enStr, i + 3, 1)))))
                        deStr = deStr & Chr(v)
                        i = i + 3
                    End If
                Else
                    deStr = deStr & c
                End If
            End If
        Else
            If c = "+" Then
                deStr = deStr & " "
            Else
                deStr = deStr & c
            End If
        End If
    Next
    URLDecode = deStr
End Function '�������

Function isvalidhex(str)
    Dim c
    isvalidhex = True
    str = UCase(str)
    If Len(str) <> 3 Then isvalidhex = False: Exit Function
    If Left(str, 1) <> "%" Then isvalidhex = False: Exit Function
    c = Mid(str, 2, 1)
    If Not (((c >= "0") And (c <= "9")) Or ((c >= "A") And (c <= "Z"))) Then isvalidhex = False: Exit Function
    c = Mid(str, 3, 1)
    If Not (((c >= "0") And (c <= "9")) Or ((c >= "A") And (c <= "Z"))) Then isvalidhex = False: Exit Function
End Function

%>

