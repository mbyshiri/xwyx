<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If CheckUserLogined() = True Then
    Call GetUser(UserName)
End If
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteName & " >> ����֧��"%></title>
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
</head>
<body>
<SCRIPT language='JavaScript1.2' src='../js/stm31.js' type='text/javascript'></SCRIPT>

<table height=114 cellSpacing=0 cellPadding=0 width=778 align=center background=../skin/Ocean/top_bg.jpg border=0>
  <tr>
    <td width=213><img src="../skin/Ocean/top_01.jpg" width="213" height="114" alt=""></td>
    <td>
      <table cellSpacing=0 cellPadding=0 width="100%" border=0>
        <tr>
          <td colSpan=2 align="right">
            <table cellSpacing=0 cellPadding=0 align=right border=0>
              <tr>
                <td><IMG height=25 src="../skin/Ocean/Announce_01.jpg" width=68></td>
                <td class=showa width=280 background=../skin/Ocean/Announce_02.jpg>&nbsp;</td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td width="83%" height=80><img src="<%= strInstallDir %>images/banner.jpg" width="468" height="60"></td>
          <td width="17%">
            <table height=89 cellSpacing=0 cellPadding=0 width=94 background=<%= strInstallDir %>Skin/images/topr.gif border=0>
              <tr>
                <td align=middle colSpan=2>
                  <table height=56 cellSpacing=0 cellPadding=0 width=79 border=0>
                    <tr>
                      <td align=middle width=26><IMG height=13 src="../skin/Ocean/arrows.gif" width=13></td>
                      <td width=68><A class=Bottom href="javascript:window.external.addFavorite('http://www.powereasy.net','��������');">�����ղ�</A></td>
                    </tr>
                    <tr>
                      <td align=middle><IMG height=13 src="../skin/Ocean/arrows.gif" width=13></td>
                      <td><A class=Bottom onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('��������');" href="http://www.powereasy.net">��Ϊ��ҳ</A></td>
                    </tr>
                    <tr>
                      <td align=middle><IMG height=13 src="../skin/Ocean/arrows.gif" width=13></td>
                      <td><A class=Bottom href="mailto:info@asp163.net">��ϵվ��</A></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>

<table width="756" border="0" align="center" cellpadding="0" cellspacing="0" class="user_border">
  <tr>
    <td valign="top">
      <table width="100%" border="0" cellpadding="5" cellspacing="0" class="user_box">
        <tr>
          <td height="200" valign='top'>
<%
Dim OrderFormID, OrderFormNum, rsOrder, dblMoneyTotal, dblMoneyReceipt, dblMoneyNeedPay
Dim PlatformID, PlatformName, AccountsID, PayGateUrl, MD5Key, Rate, PlusPoundage

OrderFormID = PE_CLng(Trim(Request("OrderFormID")))
If OrderFormID > 0 Then
    Set rsOrder = Conn.Execute("select * from PE_OrderForm where OrderFormID=" & OrderFormID & " And UserName='" & UserName & "'")
    If rsOrder.BOF And rsOrder.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ķ�����</li>"
    Else
        If rsOrder("MoneyTotal") <= rsOrder("MoneyReceipt") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ָ���Ķ����Ѿ����壬�����ٸ��</li>"
        Else
            OrderFormNum = rsOrder("OrderFormNum")
            dblMoneyTotal = rsOrder("MoneyTotal")
            dblMoneyReceipt = rsOrder("MoneyReceipt")
            dblMoneyNeedPay = dblMoneyTotal - dblMoneyReceipt
        End If
    End If
    rsOrder.Close
    Set rsOrder = Nothing
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Response.End
    End If
Else
    dblMoneyNeedPay = 100
End If

Select Case Action
Case "Step2"
    Call Step2
Case "Step3"
    Call Step3
Case Else
    Call Step1
End Select
If FoundErr = True Then
    Response.Write ErrMsg
    Response.End
End If
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>


<table cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <tr>
    <td class=menu_bottombg align=middle>
      | <a class="Bottom" href="#" onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('<%=SiteUrl%>');">��Ϊ��ҳ</a>
      | <a class="Bottom" href="javascript:window.external.addFavorite('<%=SiteUrl%>','<%=SiteName%>');">�����ղ�</a>
      | <a class="Bottom" href="mailto:<%=WebmasterEmail%>">��ϵվ��</a>
      | <a class="Bottom" href="<%=InstallDir%>FriendSite/Index.asp" target="_blank">��������</a>
      | <a class="Bottom" href="<%=InstallDir%>Copyright.asp" target="_blank">��Ȩ����</a>
      | <a class='Bottom' href='<%=InstallDir&AdminDir%>/Admin_Index.asp' target='_blank'>�����¼</a>
      |
    </td>
  </tr>
  <tr>
    <td class=bottom_bg height=80>
      <table cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
        <tr>
          <td><IMG height=80 src="<%=InstallDir%>Skin/images/bottom_left.gif" width=9></td>
          <td align=middle width="80%"> ��Ȩ���� &copy; 2003-2006</td>
          <td align=right><IMG height=80 src="<%=InstallDir%>Skin/images/bottom_r.gif" width=9></td>
        </tr>
    </table></td>
  </tr>
</table>

</body>
</html>
<%
Call CloseConn



Sub Step1()
%>
<form name='payonline' method='post'  action='PayOnline.asp'>
<table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>
  <tr>
    <td vAlign=top><table width="100%"  border="0" cellpadding="2" cellspacing="1" class="Shop_border">
        <tr>
          <td align="center" class="Shop_title"><b>�� �� ֧ �� �� ��</b>(��һ��)</td>
        </tr>
        <tr>
          <td class="Shop_tdbg">
          <table width="400" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#999999">
    <tr>
      <td colspan="2" align="center" bgcolor="#E6E6E6"><b>�� �� ֧ ��</b></td>
    </tr>
<%
    If OrderFormID > 0 Then
%>
    <tr>
      <td width="125" bgcolor="#FFFFFF">�������룺</td>
      <td width="264" bgcolor="#FFFFFF"><%=OrderFormNum%></td>
    </tr>
    <tr>
      <td width="125" bgcolor="#FFFFFF">������</td>
      <td width="264" bgcolor="#FFFFFF"><%=FormatNumber(dblMoneyTotal, 2, vbTrue, vbFalse, vbTrue)%></td>
    </tr>
    <tr>
      <td width="125" bgcolor="#FFFFFF">�� �� �</td>
      <td width="264" bgcolor="#FFFFFF"><%=FormatNumber(dblMoneyReceipt, 2, vbTrue, vbFalse, vbTrue)%></td>
    </tr>
    <tr>
      <td width="125" bgcolor="#FFFFFF">��Ҫ֧����</td>
      <td width="264" bgcolor="#FFFFFF"><%=FormatNumber(dblMoneyNeedPay, 2, vbTrue, vbFalse, vbTrue)%></td>
    </tr>
<%
    End If
    If OrderFormID = 0 Then
%>
    <tr>
      <td width="125" bgcolor="#FFFFFF">��������Ҫ��Ľ�</td>
      <td width="264" bgcolor="#FFFFFF"><input name="vMoney" type="text" id="vMoney" value="<%=FormatNumber(dblMoneyNeedPay, 2, vbTrue, vbFalse, vbTrue)%>" size="10" maxlength="20">
Ԫ</td>
    </tr>
<%
    End If
%>
    <tr>
      <td width="125" bgcolor="#FFFFFF">��ѡ������֧��ƽ̨��</td>
      <td width="264" bgcolor="#FFFFFF"><%=GetPayPlatformList%></td>
    </tr>
    <tr>
      <td colspan="2" align="center" bgcolor="#E6E6E6"><input name="OrderFormID" type="hidden" id="OrderFormID" value="<%=OrderFormID%>">
      <input name="Action" type="hidden" id="Action" value="Step2">
          <input type="submit" Name="Submit" value=" ��һ�� "></td>
    </tr>
  </table>
  </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td class=main_shadow></td>
  </tr>
</table>
</form>
<%
End Sub


Sub Step2()
    Dim vMoney
    Dim v_amount, v_mid, v_url, v_oid, v_moneytype, v_orderstatus, key_key, md5string
    Dim v_ymd, v_hms
    Dim v_ShowResultUrl
    v_ymd = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
    v_hms = Right("0" & Hour(Time), 2) & Right("0" & Minute(Time), 2) & Right("0" & Second(Time), 2)
    vMoney = Trim(Request.Form("vMoney"))

    PlatformID = Trim(Request("PlatformID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������֧��ƽ̨ID</li>"
    Else
        PlatformID = PE_CLng(PlatformID)
    End If

    If FoundErr = True Then Exit Sub

    Dim rsPayPlatform
    Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform where PlatformID=" & PlatformID & "")
    If rsPayPlatform.BOF And rsPayPlatform.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ��������֧��ƽ̨��</li>"
    Else
        If PE_CLng(rsPayPlatform("IsDisabled")) = -1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��֧��ƽ̨δ���ã�</li>"
        Else
            Select Case PlatformID
            Case 1 
                If rsPayPlatform("MD5Key") = "sldkfjkalsdjfasdf" Then 
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>Ϊ�������׵İ�ȫ,�벻Ҫʹ��ϵͳĬ�ϵ�MD5��Կ��</li>"                
                End If						          									
            Case Else
                If rsPayPlatform("MD5Key") = "aaaaaaaaaa" Then 
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>Ϊ�������׵İ�ȫ,�벻Ҫʹ��ϵͳĬ�ϵ�MD5��Կ��</li>"                
                End If			
            End Select		       					    
        End If	
        PlatformName = rsPayPlatform("ShowName")
        AccountsID = rsPayPlatform("AccountsID")
        MD5Key = rsPayPlatform("MD5Key")
        Rate = rsPayPlatform("Rate")
        PlusPoundage = rsPayPlatform("PlusPoundage")
    End If
    Set rsPayPlatform = Nothing
    
    If OrderFormID > 0 Then
        vMoney = FormatNumber(dblMoneyNeedPay, 2, vbTrue, vbFalse, vbTrue)
         'vMoney=dblMoneyTotal
    Else
        vMoney = Trim(Request("vMoney"))
    End If
    If vMoney = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����뻮���</li>"
    Else
        vMoney = Abs(PE_CDbl(vMoney))
        If vMoney < 0.01 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ÿ�λ�����ܵ���0.01Ԫ��</li>"
        Else
            If PlatformID = 11 Then
                If PlusPoundage = True Then
                    v_amount = Round(vMoney + vMoney * (Rate / 100), 1) '������
                Else
                    v_amount = vMoney
                End If
                Dim intMoney
                intMoney = Int(v_amount)
                If intMoney < v_amount Then
                    v_amount = intMoney + 1
                Else
                    v_amount = intMoney    '��Ǯ������֧�����Ϊ����ֵ
                End If
            Else
                vMoney = Round(vMoney, 2)
                If PlusPoundage = True Then
                    v_amount = Round(vMoney + vMoney * (Rate / 100), 2) '������
                Else
                    v_amount = vMoney
                End If
            End If
        End If
    End If

    If FoundErr = True Then Exit Sub


    '�õ�PaymentID
    Dim PaymentID, PaymentNum
    Dim rsPayment, sqlPayment
    Dim trs, strHiddenField
    
    PaymentNum = Prefix_PaymentNum & v_ymd & v_hms
    
    PaymentID = GetNewID("PE_Payment", "PaymentID")
    sqlPayment = "select top 1 * from PE_Payment"
    Set rsPayment = Server.CreateObject("adodb.recordset")
    rsPayment.Open sqlPayment, Conn, 1, 3
    rsPayment.AddNew
    rsPayment("PaymentID") = PaymentID
    rsPayment("UserName") = UserName
    rsPayment("OrderFormID") = OrderFormID
    rsPayment("PaymentNum") = PaymentNum
    rsPayment("eBankID") = PlatformID
    rsPayment("MoneyPay") = vMoney
    rsPayment("MoneyTrue") = v_amount
    rsPayment("PayTime") = Now()
    rsPayment("Status") = 1
    rsPayment("eBankInfo") = ""
    rsPayment("Remark") = ""
    rsPayment.Update
    rsPayment.Close
    Set rsPayment = Nothing
    

    v_mid = AccountsID
    v_moneytype = "0"               '0Ϊrmb 1Ϊdollor
    v_orderstatus = "1"             '0δ���� 1Ϊ����
    v_url = "http://" & Trim(Request.ServerVariables("HTTP_HOST")) & Trim(Request.ServerVariables("SCRIPT_NAME"))
    v_ShowResultUrl = Left(v_url, InStrRev(v_url, "/")) & "ShowResult.asp"
    v_url = Left(v_url, InStrRev(v_url, "/")) & "PayResult" & PlatformID & ".asp"

    v_oid = PaymentNum
    v_amount = FormatNumber(v_amount, 2, vbTrue, vbFalse, vbFalse)

    Select Case PlatformID
    Case 1 '��������
        PayGateUrl = "https://pay3.chinabank.com.cn/PayGate"
        v_oid = PaymentNum
        md5string = UCase(Trim(MD5(v_amount & v_moneytype & v_oid & v_mid & v_url & MD5Key, 32)))
        strHiddenField = strHiddenField & "<input type='hidden' name='v_md5info' value='" & md5string & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_mid' value='" & v_mid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_oid' value='" & v_oid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_amount' value='" & v_amount & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_moneytype'  value='" & v_moneytype & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_url' value='" & v_url & "'>" & vbCrLf
    Case 2  '�й�����֧����
        PayGateUrl = "http://www.ipay.cn/4.0/bank.shtml"
        v_oid = v_ymd & v_hms
        md5string = LCase(MD5(v_mid & v_oid & v_amount & "test@Ipay.com.cn13800138000" & MD5Key, 32))
        strHiddenField = strHiddenField & "<input type='hidden' name='v_mid' value='" & v_mid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_oid' value='" & v_oid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_amount' value='" & v_amount & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_email' value='test@Ipay.com.cn'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_mobile' value='13800138000'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_md5'    value='" & md5string & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='v_url' value='" & v_url & "'>" & vbCrLf
    Case 3  '�Ϻ���Ѹ
        PayGateUrl = "http://pay.ips.com.cn/ipayment.aspx"   '��ʽ�ӿ�
        'PayGateUrl = "http://pay.ips.net.cn/ipayment.aspx"   '���Խӿ�
        md5string = LCase(MD5(v_oid & v_amount & v_ymd & "RMB" & MD5Key, 32))
        strHiddenField = strHiddenField & "<input type='hidden' name='mer_code' value='" & v_mid & "'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='billNo' value='" & v_oid & "'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='amount' value='" & v_amount & "'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='date' value='" & v_ymd & "'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='lang' value='GB'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='Gateway_type' value='01'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='Currency_Type' value='RMB'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='Merchanturl' value='" & v_url & "'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='OrderEncodeType' value='2'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='RetEncodeType' value='12'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='RetType' value='0'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='SignMD5' value='" & md5string & "'>"
        strHiddenField = strHiddenField & "<input type='hidden' name='ServerUrl' value=''>"
    Case 4  '�й��������ݷֹ�˾
        PayGateUrl = "http://218.19.140.170/Bin/Scripts/OpenVendor/Gnete/V34/GetOvOrder.asp"

        Dim RcvCertPath, SendCertPath, SendCertPWD, MerId, OrderNo, OrderAmount, CurrCode, CallBackUrl, ResultMode
        Dim Reserved01, Reserved02, SourceText, obj, EncryptedMsg, SignedMsg, bolRet, nPayStat
        
        MerId = v_mid                           '�̻�ID����
        OrderNo = v_oid            '�̻�������
        OrderAmount = v_amount    '��������ʽ��Ԫ.�Ƿ�
        CurrCode = "CNY"          '���Ҵ��룬ֵΪ��CNY
        CallBackUrl = v_url    '֧���������URL
        ResultMode = "0"                '֧��������ط�ʽ(0-�ɹ���ʧ��֧����������أ�1-�����سɹ�֧�����)
        Reserved01 = ""                 '������1
        Reserved02 = ""                 '������2
        
        SendCertPath = "c:\certs\MERCHANT.pfx"          '���ͷ�֤��·��(�̻�֤��)
        RcvCertPath = "c:\certs\GNETEWEB-TEST.cer"          '���շ�֤��·��(����֤��)
        SendCertPWD = "12345678"      '���ͷ�֤������(�̻�֤��)
        
        '��ϳɶ���ԭʼ����
        SourceText = "MerId=" & MerId & "&" & _
                  "OrderNo=" & OrderNo & "&" & _
                  "OrderAmount=" & OrderAmount & "&" & _
                  "CurrCode=" & CurrCode & "&" & _
                  "CallBackUrl=" & CallBackUrl & "&" & _
                  "ResultMode=" & ResultMode & "&" & _
                  "Reserved01=" & Reserved01 & "&" & _
                  "Reserved02=" & Reserved02
        
        Set obj = Server.CreateObject("OpenVendorV34.NetTran")
    
        'ʹ�ý��շ�֤��Զ���ԭʼ���ݽ��м���
        If obj.EncryptMsg(SourceText, RcvCertPath) = 0 Then
            EncryptedMsg = obj.LastResult
        Else
            Response.Write obj.LastErrMsg
            Exit Sub
        End If
        
        'ʹ�÷��ͷ�֤��Զ���ԭʼ���ݽ���ǩ��
        If obj.SignMsg(SourceText, SendCertPath, SendCertPWD) = 0 Then
            SignedMsg = obj.LastResult
        Else
            Response.Write obj.LastErrMsg
            Exit Sub
        End If
        
        Set obj = Nothing
        
        
        strHiddenField = strHiddenField & "<input type='hidden' name='EncodeMsg' value='" & EncryptedMsg & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='SignMsg' value='" & SignedMsg & "'>" & vbCrLf
    Case 5  '����֧��
        PayGateUrl = "http://www.yeepay.com/Pay/WestPayReceiveOrderFromMerchant.asp"
        strHiddenField = strHiddenField & "<input type='hidden' name='MerchantID' value='" & v_mid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='OrderNumber' value='" & v_oid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='OrderAmount' value='" & v_amount & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='PostBackURL' value='" & v_url & "'>" & vbCrLf
    Case 6   '�׸�ͨ
        PayGateUrl = "http://pay.xpay.cn/Pay.aspx"
        md5string = LCase(MD5(MD5Key & ":" & v_amount & "," & v_oid & "," & v_mid & ",bank,,sell,,2.0", 32))
        
        strHiddenField = strHiddenField & "<input type='hidden' name='Tid' value='" & v_mid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='Bid' value='" & v_oid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='Prc' value='" & v_amount & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='url' value='" & v_url & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='Card' value='bank'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='Scard' value=''>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='ActionCode' value='sell'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='ActionParameter' value=''>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='Ver' value='2.0'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='Pdt' value='" & SiteName & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='type' value=''>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='lang' value='gb2312'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='md' value='" & md5string & "'>" & vbCrLf
    Case 7   '����֧��
        PayGateUrl = "https://www.cncard.net/purchase/getorder.asp"
        md5string = LCase(MD5(v_mid & v_oid & v_amount & v_ymd & "01" & v_url & "00" & MD5Key, 32))
        
        strHiddenField = strHiddenField & "<input type='hidden' name='c_mid' value='" & v_mid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_order' value='" & v_oid & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_orderamount' value='" & v_amount & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_ymd' value='" & v_ymd & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_moneytype' value='0'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_retflag' value='1'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_paygate' value=''>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_returl' value='" & v_url & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_memo1' value=''>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_memo2' value=''>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_language' value='0'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='notifytype' value='0'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='c_signstr' value='" & md5string & "'>" & vbCrLf
    Case 8, 12 '֧����
        v_oid = PaymentNum

        Dim InstantPay
        Dim Partner
        Dim ArrMD5Key
        If InStr(MD5Key, "|") > 0 Then
            ArrMD5Key = Split(MD5Key, "|")
            If UBound(ArrMD5Key) = 1 Then
                Partner = ArrMD5Key(1)
                MD5Key = ArrMD5Key(0)
            End If
        End If
                
        v_ShowResultUrl = v_ShowResultUrl & "?PayMessage=ok"
        If PlatformID = 12 Then '֧������ʱ����
            PayGateUrl = "https://www.alipay.com/cooperate/gateway.do"
            Dim myString
            myString = "discount=0" & "&notify_url=" & v_url & "&out_trade_no=" & v_oid & "&partner=" & Partner & "&payment_type=1" & "&price=" & v_amount & "&quantity=1" & "&return_url=" & v_ShowResultUrl & "&seller_email=" & v_mid & "&service=create_direct_pay_by_user&subject=" & v_oid & MD5Key
            md5string = LCase(MD5(myString, 32))
            strHiddenField = strHiddenField & "<input type='hidden' name='discount' value='0'>" '��Ʒ�ۿ�
            strHiddenField = strHiddenField & "<input type='hidden' name='notify_url' value='" & v_url & "'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='out_trade_no' value='" & v_oid & "'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='payment_type' value='1'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='partner' value='" & Partner & "'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='price' value='" & v_amount & "'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='quantity' value='1'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='seller_email' value='" & v_mid & "'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='service' value='create_direct_pay_by_user'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='subject' value='" & v_oid & "'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='sign' value='" & md5string & "'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='sign_type' value='MD5'>"
            strHiddenField = strHiddenField & "<input type='hidden' name='return_url' value='" & v_ShowResultUrl & "'>"
        Else
            '��������������Ʒ�����
            Dim rsOrderItem
            Dim IsFabrication
            Dim transport
            IsFabrication = False
            If OrderFormID = 0 Then
                IsFabrication = True '��Ա��ֵ,��Ϊ������Ʒ
            Else
                Set rsOrderItem = Conn.Execute("select I.ItemID,P.ProductID,P.ProductName,P.ProductKind,I.Amount from PE_OrderFormItem I inner join PE_Product P on I.ProductID=P.ProductID where I.OrderFormID=" & OrderFormID & " order by I.ItemID")
                Do While Not rsOrderItem.EOF
                    If rsOrderItem("ProductKind") = 3 Then
                        IsFabrication = True
                    Else
                        IsFabrication = False
                        Exit Do
                    End If
                    rsOrderItem.MoveNext
                Loop
            End If
            If Partner = "" Then   '�ɽӿ�
                PayGateUrl = "https://www.alipay.com/payto:" & v_mid
                If IsFabrication Then
                    transport = 3
                    md5string = LCase((MD5("cmd" & "0001" & "subject" & v_oid & "order_no" & v_oid & "price" & v_amount & "transport" & transport & "seller" & v_mid & "partner" & "2088001048757497" & MD5Key, 32)))
                Else
                    md5string = LCase((MD5("cmd" & "0001" & "subject" & v_oid & "order_no" & v_oid & "price" & v_amount & "seller" & v_mid & "partner" & "2088001048757497" & MD5Key, 32)))
                End If
           
                strHiddenField = strHiddenField & "<input type='hidden' name='cmd' value='0001'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='subject' value='" & v_oid & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='order_no' value='" & v_oid & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='price' value='" & v_amount & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='partner' value='2088001048757497'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='ac'  value='" & md5string & "'>" & vbCrLf
                If IsFabrication Then strHiddenField = strHiddenField & "<input type='hidden' name='transport'  value='3'>" & vbCrLf
            Else   '�½ӿ�
                PayGateUrl = "https://www.alipay.com/cooperate/gateway.do"
                If IsFabrication Then
                    md5string = LCase(MD5("notify_url=" & v_url & "&out_trade_no=" & v_oid & "&partner=" & Partner & "&price=" & v_amount & "&quantity=1" & "&return_url=" & v_ShowResultUrl & "&seller_email=" & v_mid & "&service=create_digital_goods_trade_p&subject=" & v_oid & MD5Key, 32))
                Else
                    md5string = LCase(MD5("logistics_fee=0&logistics_payment=SELLER_PAY&logistics_type=EXPRESS&notify_url=" & v_url & "&out_trade_no=" & v_oid & "&partner=" & Partner & "&payment_type=1&price=" & v_amount & "&quantity=1" & "&return_url=" & v_ShowResultUrl & "&seller_email=" & v_mid & "&service=trade_create_by_buyer&subject=" & v_oid & MD5Key, 32))
                End If
                               
                If IsFabrication Then
                    strHiddenField = strHiddenField & "<input type='hidden' name='service' value='create_digital_goods_trade_p'>" & vbCrLf
                Else
                    strHiddenField = strHiddenField & "<input type='hidden' name='service' value='trade_create_by_buyer'>" & vbCrLf
                    strHiddenField = strHiddenField & "<input type='hidden' name='logistics_type' value='EXPRESS'>" & vbCrLf
                    strHiddenField = strHiddenField & "<input type='hidden' name='logistics_fee' value='0'>" & vbCrLf
                    strHiddenField = strHiddenField & "<input type='hidden' name='logistics_payment' value='SELLER_PAY'>" & vbCrLf
                    strHiddenField = strHiddenField & "<input type='hidden' name='payment_type' value='1'>" & vbCrLf
                End If
                strHiddenField = strHiddenField & "<input type='hidden' name='seller_email' value='" & v_mid & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='subject' value='" & v_oid & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='out_trade_no' value='" & v_oid & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='price' value='" & v_amount & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='partner' value='" & Partner & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='quantity' value='1'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='notify_url' value='" & v_url & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='return_url' value='" & v_ShowResultUrl & "'>"
                strHiddenField = strHiddenField & "<input type='hidden' name='sign' value='" & md5string & "'>" & vbCrLf
                strHiddenField = strHiddenField & "<input type='hidden' name='sign_type' value='MD5'>" & vbCrLf
            End If
        End If
    Case 9 '��Ǯ֧��2.0�ӿ�
        PayGateUrl = "https://www.99bill.com/gateway/recvMerchantInfoAction.htm"
        Dim merchantAcctId, key, inputCharset, pageUrl, bgUrl, version, language, signType, payerName, payerContactType, payerContact
        Dim orderTime, productName, productNum, productId, productDesc, ext1, ext2, payType, bankId, redoFlag, pid, signMsgVal, orderId
        merchantAcctId = v_mid   '�����˻���
        key = MD5Key '������Կ
        inputCharset = "3" '1����UTF-8; 2����GBK; 3����gb2312
        pageUrl = v_url '����֧�������ҳ���ַ
        bgUrl = "" '����������֧������ĺ�̨��ַ
        version = "v2.0" '���ذ汾.�̶�ֵ
        language = "1" '1�������ģ�2����Ӣ��
        signType = "1" '1����MD5ǩ��
        payerName = "" '֧��������
        payerContactType = "" '֧������ϵ��ʽ���� 1����Email��2�����ֻ���
        payerContact = "" '֧������ϵ��ʽ,ֻ��ѡ��Email���ֻ���
        orderId = v_oid '�̻�������
        OrderAmount = v_amount * 100 '�������,�Է�Ϊ��λ
        orderTime = v_ymd & v_hms '�����ύʱ��,14λ����
        productName = "" '��Ʒ����
        productNum = "" '��Ʒ����
        productId = "" '��Ʒ����
        productDesc = "" '��Ʒ����
        ext1 = "" '��չ�ֶ�1,��֧��������ԭ�����ظ��̻�
        ext2 = "" '��չ�ֶ�2
        payType = "00" '֧����ʽ,00�����֧��,��ʾ��Ǯ֧�ֵĸ���֧����ʽ,11���绰����֧��,12����Ǯ�˻�֧��,13������֧��,14��B2B֧��
        bankId = "" '���д���,ʵ��ֱ����ת������ҳ��ȥ֧��,�������μ� �ӿ��ĵ����д����б�,ֻ��payType=10ʱ�������ò���
        redoFlag = "1" 'ͬһ������ֹ�ظ��ύ��־:1����ͬһ������ֻ�����ύ1��,0��ʾͬһ��������û��֧���ɹ���ǰ���¿��ظ��ύ���
        pid = "" '��Ǯ�ĺ��������˻���

        signMsgVal = appendParam(signMsgVal, "inputCharset", inputCharset)
        signMsgVal = appendParam(signMsgVal, "pageUrl", pageUrl)
        signMsgVal = appendParam(signMsgVal, "bgUrl", bgUrl)
        signMsgVal = appendParam(signMsgVal, "version", version)
        signMsgVal = appendParam(signMsgVal, "language", language)
        signMsgVal = appendParam(signMsgVal, "signType", signType)
        signMsgVal = appendParam(signMsgVal, "merchantAcctId", merchantAcctId)
        signMsgVal = appendParam(signMsgVal, "payerName", payerName)
        signMsgVal = appendParam(signMsgVal, "payerContactType", payerContactType)
        signMsgVal = appendParam(signMsgVal, "payerContact", payerContact)
        signMsgVal = appendParam(signMsgVal, "orderId", v_oid)
        signMsgVal = appendParam(signMsgVal, "orderAmount", OrderAmount)
        signMsgVal = appendParam(signMsgVal, "orderTime", orderTime)
        signMsgVal = appendParam(signMsgVal, "productName", productName)
        signMsgVal = appendParam(signMsgVal, "productNum", productNum)
        signMsgVal = appendParam(signMsgVal, "productId", productId)
        signMsgVal = appendParam(signMsgVal, "productDesc", productDesc)
        signMsgVal = appendParam(signMsgVal, "ext1", ext1)
        signMsgVal = appendParam(signMsgVal, "ext2", ext2)
        signMsgVal = appendParam(signMsgVal, "payType", payType)
        signMsgVal = appendParam(signMsgVal, "bankId", bankId)
        signMsgVal = appendParam(signMsgVal, "redoFlag", redoFlag)
        signMsgVal = appendParam(signMsgVal, "pid", pid)
        signMsgVal = appendParam(signMsgVal, "key", key)
        md5string = UCase(MD5(signMsgVal, 32))
        strHiddenField = strHiddenField & "<input type='hidden' name='inputCharset' value='" & inputCharset & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='bgUrl' value='" & bgUrl & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='pageUrl' value='" & pageUrl & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='version' value='" & version & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='language' value='" & language & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='signType' value='" & signType & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='signMsg' value='" & md5string & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='merchantAcctId' value='" & merchantAcctId & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payerName' value='" & payerName & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payerContactType' value='" & payerContactType & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payerContact' value='" & payerContact & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='orderId' value='" & orderId & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='orderAmount' value='" & OrderAmount & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='orderTime' value='" & orderTime & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productName' value='" & productName & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productNum' value='" & productNum & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productId' value='" & productId & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productDesc' value='" & productDesc & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='ext1' value='" & ext1 & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='ext2' value='" & ext2 & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payType' value='" & payType & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='bankId' value='" & bankId & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='redoFlag' value='" & redoFlag & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='pid' value='" & pid & "'>" & vbCrLf
    Case 11 '��Ǯ������
        PayGateUrl = "https://www.99bill.com/szxgateway/recvMerchantInfoAction.htm"
        Dim cardNumber, cardPwd, fullAmountFlag

        merchantAcctId = v_mid '�����������˻���
        key = MD5Key '���������������Կ
        inputCharset = "3" '1����UTF-8; 2����GBK; 3����gb2312
        bgUrl = "" '����������֧������ĺ�̨��ַ
        pageUrl = v_url '����֧�������ҳ���ַ
        version = "v2.0" '���ذ汾.�̶�ֵ
        language = "1" '1�������ģ�2����Ӣ��
        signType = "1" 'ǩ������.�̶�ֵ
        payerName = "" '֧��������
        payerContactType = "" '֧������ϵ��ʽ����,1����Email��2�����ֻ���
        payerContact = "" '֧������ϵ��ʽ,ֻ��ѡ��Email���ֻ���
        orderId = v_oid '�̻�������
        OrderAmount = v_amount * 100 '�������,�Է�Ϊ��λ����������������
        orderTime = v_ymd & v_hms '�����ύʱ��
        productName = "" '��Ʒ����
        productNum = "" '��Ʒ����
        productId = "" '��Ʒ����
        cardNumber = "" '�����п����,�����̻������������п���ֱ������ʱ��д
        productDesc = "" '��Ʒ����
        ext1 = "" '��չ�ֶ�1
        ext2 = "" '��չ�ֶ�2
        payType = "00" 'ֻ��ѡ��00,����֧�������п��Ϳ�Ǯ�ʻ�֧��
        cardPwd = "" '�����п�����,�����̻������������п���ֱ������ʱ��д

        'ȫ��֧����־       ''0�������С�ڶ������ʱ����֧�����Ϊʧ�ܣ�1�������С�ڶ�������Ƿ���֧�����Ϊ�ɹ���ͬʱ��������ʵ��֧����Ϊ�����п������.����̻����������п���ֱ��ʱ���������̶�ֵΪ1
        fullAmountFlag = "0" '0�������С�ڶ������ʱ����֧�����Ϊʧ��

        ''����ذ�������˳��͹�����ɼ��ܴ���
        signMsgVal = appendParam(signMsgVal, "inputCharset", inputCharset)
        signMsgVal = appendParam(signMsgVal, "bgUrl", bgUrl)
        signMsgVal = appendParam(signMsgVal, "pageUrl", pageUrl)
        signMsgVal = appendParam(signMsgVal, "version", version)
        signMsgVal = appendParam(signMsgVal, "language", language)
        signMsgVal = appendParam(signMsgVal, "signType", signType)
        signMsgVal = appendParam(signMsgVal, "merchantAcctId", merchantAcctId)
        signMsgVal = appendParam(signMsgVal, "payerName", payerName)
        signMsgVal = appendParam(signMsgVal, "payerContactType", payerContactType)
        signMsgVal = appendParam(signMsgVal, "payerContact", payerContact)
        signMsgVal = appendParam(signMsgVal, "orderId", orderId)
        signMsgVal = appendParam(signMsgVal, "orderAmount", OrderAmount)
        signMsgVal = appendParam(signMsgVal, "payType", payType)
        signMsgVal = appendParam(signMsgVal, "cardNumber", cardNumber)
        signMsgVal = appendParam(signMsgVal, "cardPwd", cardPwd)
        signMsgVal = appendParam(signMsgVal, "fullAmountFlag", fullAmountFlag)
        signMsgVal = appendParam(signMsgVal, "orderTime", orderTime)
        signMsgVal = appendParam(signMsgVal, "productName", productName)
        signMsgVal = appendParam(signMsgVal, "productNum", productNum)
        signMsgVal = appendParam(signMsgVal, "productId", productId)
        signMsgVal = appendParam(signMsgVal, "productDesc", productDesc)
        signMsgVal = appendParam(signMsgVal, "ext1", ext1)
        signMsgVal = appendParam(signMsgVal, "ext2", ext2)
        signMsgVal = appendParam(signMsgVal, "key", key)
        md5string = UCase(MD5(signMsgVal, 32))

        
        strHiddenField = strHiddenField & "<input type='hidden' name='inputCharset' value='" & inputCharset & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='bgUrl' value='" & bgUrl & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='pageUrl' value='" & pageUrl & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='version' value='" & version & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='language' value='" & language & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='signType' value='" & signType & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='merchantAcctId' value='" & merchantAcctId & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payerName' value='" & payerName & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payerContactType' value='" & payerContactType & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payerContact' value='" & payerContact & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='orderId' value='" & orderId & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='orderAmount' value='" & OrderAmount & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='orderTime' value='" & orderTime & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productName' value='" & productName & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productNum' value='" & productNum & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productId' value='" & productId & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='productDesc' value='" & productDesc & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='ext1' value='" & ext1 & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='ext2' value='" & ext2 & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='payType' value='" & payType & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='fullAmountFlag' value='" & fullAmountFlag & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='cardNumber' value='" & cardNumber & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='cardPwd' value='" & cardPwd & "'>" & vbCrLf
        strHiddenField = strHiddenField & "<input type='hidden' name='signMsg' value='" & md5string & "'>" & vbCrLf
    Case 13  '��Ѷ�Ƹ�ͨ
        Dim transaction_id
        transaction_id = v_mid & v_ymd & Right(v_oid, 10)
        PayGateUrl = "https://www.tenpay.com/cgi-bin/v1.0/pay_gate.cgi"
        md5string = UCase(MD5("cmdno=1&date=" & v_ymd & "&bargainor_id=" & v_mid & "&transaction_id=" & transaction_id & "&sp_billno=" & v_oid & "&total_fee=" & v_amount * 100 & "&fee_type=1&return_url=" & v_url & "&attach=my_magic_string&key=" & MD5Key, 32))
        strHiddenField = strHiddenField & "<input type='hidden' name='cmdno' value='1'>"   'ҵ�����,1��ʾ֧��
        strHiddenField = strHiddenField & "<input type='hidden' name='date' value='" & v_ymd & "'>"   '�̻�����
        strHiddenField = strHiddenField & "<input type='hidden' name='bank_type' value='0'>"  '��������:�Ƹ�ͨ,0
        strHiddenField = strHiddenField & "<input type='hidden' name='desc' value='" & v_oid & "'>"    '���׵���Ʒ����
        strHiddenField = strHiddenField & "<input type='hidden' name='purchaser_id' value=''>"   '�û�(��)�ĲƸ�ͨ�ʻ�,����Ϊ��
        strHiddenField = strHiddenField & "<input type='hidden' name='bargainor_id' value='" & v_mid & "'>"  '�̼ҵ��̻���
        strHiddenField = strHiddenField & "<input type='hidden' name='transaction_id' value='" & transaction_id & "'>"   '���׺�(������)
        strHiddenField = strHiddenField & "<input type='hidden' name='sp_billno' value='" & PaymentNum & "'>"  '�̻�ϵͳ�ڲ��Ķ�����
        strHiddenField = strHiddenField & "<input type='hidden' name='total_fee' value='" & v_amount * 100 & "'>" '�ܽ��Է�Ϊ��λ
        strHiddenField = strHiddenField & "<input type='hidden' name='fee_type' value='1'>"  '�ֽ�֧������,1�����
        strHiddenField = strHiddenField & "<input type='hidden' name='return_url' value='" & v_url & "'>" '���ղƸ�ͨ���ؽ����URL
        strHiddenField = strHiddenField & "<input type='hidden' name='attach' value='my_magic_string'>" '�̼����ݰ���ԭ������
        strHiddenField = strHiddenField & "<input type='hidden' name='sign' value='" & md5string & "'>" 'MD5ǩ��
    End Select
%>
<form name='payonline' method='post' action='<%=PayGateUrl%>'>
<table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>
  <tr>
    <td vAlign=top><table width="100%"  border="0" cellpadding="2" cellspacing="1" class="Shop_border">
        <tr>
          <td align="center" class="Shop_title"><b>�� �� ֧ �� �� ��</b>(�ڶ���)</td>
        </tr>
        <tr>
          <td class="Shop_tdbg">
            <table width="400" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
              <tr>
                <td colspan="2" align="center" bgcolor="#E6E6E6"><b>ȷ �� �� ��</b></td>
              </tr>
              <tr>
                <td width="100" align="right" bgcolor="#FFFFFF">֧�����кţ�</td>
                <td width="289" align="center" bgcolor="#FFFFFF"><%=PaymentNum%></td>
              </tr>
              <tr>
                <td width="100" align="right" bgcolor="#FFFFFF">֧����</td>
                <td align="center" bgcolor="#FFFFFF">��<%=FormatNumber(vMoney, 2, vbTrue, vbFalse, vbTrue)%></td>
              </tr>
              <tr>
                <td width="100" align="right" bgcolor="#FFFFFF">�����ѣ�</td>
                <td align="center" bgcolor="#FFFFFF"><%=FormatNumber(Rate, 2, vbTrue, vbFalse, vbTrue) & "%"%></td>
              </tr>
              <tr>
                <td width="100" align="right" bgcolor="#FFFFFF">ʵ�ʻ����</td>
                <td align="center" bgcolor="#FFFFFF">��<%=v_amount%></td>
              </tr>
              <tr bgcolor="#E6E6E6">
                <td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;�����ȷ��֧������ť�󣬽�����<%=PlatformName%>֧�����棬�ڴ�ҳ��ѡ���������п���</td>
              </tr>
              <tr align="center" bgcolor="#E6E6E6">
                <td colspan="2"><input type="submit" id="Submit" value=" ȷ��֧�� ">&nbsp;
      <input type="button" name="Submit" value=" ȡ��֧�� " onclick="window.location.href='../User/User_Payment.asp?Action=Cancel&PaymentID=<%=PaymentID%>'">
                <%=strHiddenField%></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td class=main_shadow></td>
  </tr>
</table>
</form>
<%
End Sub

Function GetPayPlatformList()
    Dim rsPayPlatform, strList
    Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform where IsDisabled=" & PE_False & " order by OrderID asc")
    If rsPayPlatform.BOF And rsPayPlatform.EOF Then
        strList = "û�������κ�����֧��ƽ̨"
    Else
        Do While Not rsPayPlatform.EOF
            strList = strList & "<input type='radio' Name='PlatformID' value='" & rsPayPlatform("PlatformID") & "'"
            If rsPayPlatform("IsDefault") = True Then strList = strList & "checked"
            strList = strList & ">" & rsPayPlatform("ShowName") & "<br>"
            If rsPayPlatform("Description") <> "" Then strList = strList & rsPayPlatform("Description") & "<br>"


            rsPayPlatform.MoveNext
        Loop
    End If
    Set rsPayPlatform = Nothing
    GetPayPlatformList = strList
End Function

'������ֵ��Ϊ�յĲ�������ַ���(��Ǯ)
Function appendParam(returnStr, paramId, paramValue)
    If returnStr <> "" Then
        If paramValue <> "" Then
            returnStr=returnStr&"&"&paramId&"="&paramValue
        End If
    Else
        If paramValue <> "" Then
            returnStr=paramId&"="&paramValue
        End If
    End If
    appendParam = returnStr
End Function
%>




