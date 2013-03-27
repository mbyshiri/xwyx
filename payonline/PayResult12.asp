<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5_New.asp"-->
<!--#include file="UpdateOrder.asp"-->
<HTML>
<HEAD>
<TITLE>在线支付结果</TITLE>
</HEAD>
<BODY style="font-size:9pt;">
<%
Const IsMessageShow = True
Const PlatformID = 12  '支付宝即时到帐
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

Dim trade_status, sign, MySign, Retrieval
Dim mystr, Count, i, minmax, minmaxSlot, j, mark, temp, value, md5str, notify_id

v_oid = DelStr(Request("out_trade_no"))            '商户定单号
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

            
'获取POST过来的参数
mystr = Split(URLDecode(Request.Form), "&")
Count = UBound(mystr)

'对参数排序
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

'构造md5摘要字符串
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
'生成md5摘要
MySign = PE_Md5.MD5(md5str)

returnTxt = "fail"
If sign = MySign And ResponseTxt = "true" Then
   If trade_status = "TRADE_FINISHED" Then
     '付款成功,更新数据库
        returnTxt = "success"
        PaySuccess = True                '交易成功，更新订单
    Else
        returnTxt = "success"
    End If
End If
Response.Write returnTxt

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
    DelStr = Replace(DelStr, "　", "")
    DelStr = Replace(DelStr, "%20", "")
    DelStr = Replace(DelStr, "--", "")
    DelStr = Replace(DelStr, "==", "")
    DelStr = Replace(DelStr, "<", "")
    DelStr = Replace(DelStr, ">", "")
    DelStr = Replace(DelStr, "%", "")
End Function

'取得请求后返回的html Stream
Function GetBody(strURL)
    On Error Resume Next
    Dim Retrieval
    Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
    Retrieval.Open "GET", strURL, False, "", ""
    Retrieval.Send
    GetBody = Retrieval.ResponseText
    Set Retrieval = Nothing
End Function

'对post传递过来的参数作urldecode编码处理(支付宝，新接口)
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
End Function '处理完毕

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
</BODY>
</HTML>

