<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If EnableLinkReg <> True Then
    FoundErr = True
    ErrMsg = ErrMsg & "<br><li>����Աû�п��������������룡</li>"
Else
    If Action = "Modify" Then
        Call SaveModify
    Else
        Call main
    End If
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Call CloseConn

Sub main()
    Dim ID, rsLink, sqlLink
    ID = PE_CLng(Trim(request("ID")))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��ָ����������ID</li>"
        Exit Sub
    End If
    sqlLink = "select * from PE_FriendSite where Passed=" & PE_True & " and ID=" & ID
    Set rsLink = Conn.Execute(sqlLink)
    If rsLink.bof And rsLink.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Ҳ����������ӻ�����������δ���ͨ����</li>"
        rsLink.Close
        Set rsLink = Nothing
        Exit Sub
    End If
%>
<html>
<head>
<title>�޸���������</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
<script language = "JavaScript">
<!--
function CheckForm(){
  if(document.myform.SiteName.value==""){
    alert("��������վ���ƣ�");
    document.myform.SiteName.focus();
    return false;
  }
  if(document.myform.SiteUrl.value=="" || document.myform.SiteUrl.value=="http://"){
    alert("��������վ��ַ��");
    document.myform.SiteUrl.focus();
    return false;
  }
  if(document.myform.SiteAdmin.value==""){
    alert("������վ��������");
    document.myform.SiteAdmin.focus();
    return false;
  }
  if(document.myform.OldSitePassword.value==""){
    alert("������ԭ�����룡");
    document.myform.OldSitePassword.focus();
    return false;
  }
  if(document.myform.SitePwdConfirm.value!=""||document.myform.SitePassword.value!=""){
    if(document.myform.SitePwdConfirm.value!=document.myform.SitePassword.value){
      alert("��վ������ȷ�����벻һ�£�");
      document.myform.SitePwdConfirm.focus();
      document.myform.SitePwdConfirm.select();
      return false;
    }
  }
  if(document.myform.SiteIntro.value==""){
    alert("��������վ��飡");
    document.myform.SiteIntro.focus();
    return false;
  }
}
//-->
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="myform" onsubmit="return CheckForm()" action="FriendSiteModify.asp">
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="center_tdbgall">
    <tr>
      <td align="center"><br>
        <table width="400" border="0" cellspacing="0" cellpadding="0" class="main_title_575">
          <tr>
            <td><b>�޸�����������Ϣ</b></td>
          </tr>
        </table>
        <table border="0" cellpadding="2" cellspacing="1" width="400" class="main_tdbg_575">
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">�������</td>
            <td height="25"><select name="KindID" id="KindID">
                <%=GetFsKind_Option(1, rsLink("KindID"))%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">����ר�⣺</td>
            <td height="25"><select name="SpecialID" id="SpecialID">
                <%=GetFsKind_Option(2, rsLink("SpecialID"))%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">��վ���ƣ�</td>
            <td height="25"><input name="SiteName" size="30" maxlength="20" title="����������������վ���ƣ����Ϊ20������" value="<%=rsLink("SiteName")%>">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">��վ��ַ��</td>
            <td height="25"><input name="SiteUrl" size="30" maxlength="100" type="text" value="<%=rsLink("SiteUrl")%>" title="����������������վ��ַ�����Ϊ50���ַ���ǰ������http://">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">��վLogo��</td>
            <td height="25"><input name="LogoUrl" size="30" maxlength="100" type="text" value="<%=rsLink("LogoUrl")%>" title="����������������վLogoUrl��ַ�����Ϊ50���ַ���������ڵ�һѡ��ѡ������������ӣ�����Ͳ�����"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">վ��������</td>
            <td height="25"><input name="SiteAdmin" size="30" maxlength="20" type="text" title="�������������Ĵ����ˣ���Ȼ��֪������˭�������Ϊ20���ַ�" value="<%=rsLink("SiteAdmin")%>">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">�����ʼ���</td>
            <td height="25"><input name="SiteEmail" size="30" maxlength="30" type="text" value="<%=rsLink("SiteEmail")%>" title="����������������ϵ�����ʼ������Ϊ30���ַ�"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">ԭ�����룺</td>
            <td height="25"><input name="OldSitePassword" type="password" id="OldSitePassword" size="20" maxlength="20">
              <font color="#FF0000">* ��������</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">�������룺</td>
            <td height="25"><input name="SitePassword" type="password" id="SitePassword" size="20" maxlength="20">
              <font color="#0000FF">�����޸ģ��뱣��Ϊ��</font></td>
          </tr>
          <tr class="tdbg">
            <td height="25" align="right">ȷ�����룺</td>
            <td height="25"><input name="SitePwdConfirm" type="password" id="SitePwdConfirm" size="20" maxlength="20"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" align="right">��վ��飺</td>
            <td valign="middle"><textarea name="SiteIntro" cols="40" rows="5" id="SiteIntro" title="����������������վ�ļ򵥽���"><%=PE_ConvertBR(rsLink("SiteIntro"))%></textarea></td>
          </tr>
          <tr class="tdbg">
            <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Modify"><input type="submit" value=" ȷ �� " name="cmdOk">
              <input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>"></td>
          </tr>
        </table>
        <br>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
<%
    rsLink.Close
    Set rsLink = Nothing
End Sub

Sub SaveModify()
    Dim ID, KindID, SpecialID, LinkType, LinkSiteName, LinkSiteUrl, LinkLogoUrl, LinkSiteAdmin, LinkSiteEmail, OldSitePassword, LinkSitePassword, LinkSitePwdConfirm, LinkSiteIntro
    ID = PE_CLng(Trim(request.Form("ID")))
    KindID = PE_CLng(Trim(request.Form("KindID")))
    SpecialID = PE_CLng(Trim(request.Form("SpecialID")))
    LinkSiteName = Trim(request("SiteName"))
    LinkSiteUrl = Trim(request("SiteUrl"))
    LinkLogoUrl = Trim(request("LogoUrl"))
    LinkSiteAdmin = Trim(request("SiteAdmin"))
    LinkSiteEmail = Trim(request("SiteEmail"))
    OldSitePassword = Trim(request("OldSitePassword"))
    LinkSitePassword = Trim(request("SitePassword"))
    LinkSitePwdConfirm = Trim(request("SitePwdConfirm"))
    LinkSiteIntro = Trim(request("SiteIntro"))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>����ȷ����������ID</li>"
    End If
    If LinkSiteName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ���Ʋ���Ϊ�գ�</li>"
    End If
    If LinkSiteUrl = "" Or LinkSiteUrl = "http://" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ��ַ����Ϊ�գ�</li>"
    End If
    If LinkSiteAdmin = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>վ����������Ϊ�գ�</li>"
    End If
    If LinkSiteEmail <> "" And IsValidEmail(LinkSiteEmail) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�����ʼ���ַ����</li>"
    End If
    If OldSitePassword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վԭ�����벻��Ϊ�գ�</li>"
    End If
    If LinkSitePwdConfirm <> LinkSitePassword Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>����վ������ȷ�����벻һ�£�</li>"
    End If
    If LinkSiteIntro = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ��鲻��Ϊ�գ�</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If LinkLogoUrl = "" Or LinkLogoUrl = "http://" Then
        LinkType = 2
    Else
        LinkType = 1
    End If

    Dim sqlLink, rsLink
    sqlLink = "select * from PE_FriendSite where ID=" & ID
    Set rsLink = Server.CreateObject("Adodb.RecordSet")
    rsLink.open sqlLink, Conn, 1, 3
    If rsLink.bof And rsLink.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Ҳ���ָ�����������ӣ�</li>"
    Else
        If MD5(OldSitePassword, 16) <> rsLink("SitePassword") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>������ľ���վ���벻�ԣ�û��Ȩ���޸ģ�</li>"
            rsLink.Close
            Set rsLink = Nothing
            Exit Sub
        End If
        rsLink("KindID") = KindID
        rsLink("SpecialID") = SpecialID
        rsLink("LinkType") = LinkType
        rsLink("SiteName") = ReplaceBadChar(LinkSiteName)
        rsLink("SiteUrl") = ReplaceUrlBadChar(LinkSiteUrl)
        rsLink("LogoUrl") = ReplaceUrlBadChar(LinkLogoUrl)
        rsLink("SiteAdmin") = PE_HTMLEncode(LinkSiteAdmin)
        rsLink("SiteEmail") = PE_HTMLEncode(LinkSiteEmail)
        If LinkSitePassword <> "" Then
            rsLink("SitePassword") = MD5(LinkSitePassword, 16)
        End If
        rsLink("SiteIntro") = PE_HTMLEncode(LinkSiteIntro)
        rsLink("UpdateTime") = Now
        rsLink("Passed") = False
        rsLink.Update
        Call WriteSuccessMsg("�޸��������ӳɹ�����ȴ�����Ա���ͨ����", ComeUrl)
    End If
    rsLink.Close
    Set rsLink = Nothing
End Sub

Function GetFsKind_Option(iKindType, KindID)
    Dim sqlFsKind, rsFsKind, strOption
    strOption = "<option value='0'"
    If KindID = "" Then
        strOption = strOption & " selected"
    End If
    If iKindType = 1 Then
        strOption = strOption & ">�������κ����</option>"
    ElseIf iKindType = 2 Then
        strOption = strOption & ">�������κ�ר��</option>"
    End If
    sqlFsKind = "select * from PE_FsKind"
    If iKindType > 0 Then
        sqlFsKind = sqlFsKind & " where KindType=" & iKindType
    End If
    sqlFsKind = sqlFsKind & " order by KindID"
    Set rsFsKind = Conn.Execute(sqlFsKind)
    Do While Not rsFsKind.EOF
        If rsFsKind("KindID") = KindID Then
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "' selected>" & rsFsKind("KindName") & "</option>"
        Else
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "'>" & rsFsKind("KindName") & "</option>"
        End If
        rsFsKind.movenext
    Loop
    rsFsKind.Close
    Set rsFsKind = Nothing
    GetFsKind_Option = strOption
End Function

%>
