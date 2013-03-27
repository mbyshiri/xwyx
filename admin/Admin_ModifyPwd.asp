<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "ModifyPwd"   '其他权限

Response.Write "<html><head><title>修改管理员信息</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf

'执行的操作
Select Case Action
Case "Modify"
    Call ModifyPwd
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Response.Write "<br><br>" & vbCrLf
    Response.Write "<script language='JavaScript'>" & vbCrLf
    Response.Write "function CheckForm()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(document.myform.Password.value=='')" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "      alert('密码不能为空！');" & vbCrLf
    Response.Write "      document.myform.Password.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if((document.myform.Password.value)!=(document.myform.PwdConfirm.value))" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "      alert('初始密码与确认密码不同！');" & vbCrLf
    Response.Write "      document.myform.PwdConfirm.select();" & vbCrLf
    Response.Write "      document.myform.PwdConfirm.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_ModifyPwd.asp'>" & vbCrLf
    Response.Write "  <table width='300' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='2' align='center'><strong>修 改 管 理 员 密 码</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100' align='right'><strong>会 员 名：</strong></td>" & vbCrLf
    Response.Write "      <td>" & AdminName & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100' align='right'><strong>会员权限：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Select Case AdminPurview
    Case 1
        Response.Write "超级管理员"
    Case 2
        Response.Write "普通管理员"
    End Select
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100' align='right'><strong>新 密 码：</strong></td>" & vbCrLf
    Response.Write "      <td><input type='password' name='Password' onkeyup='javascript:EvalPwdStrength(document.forms[0],this.value);' onmouseout='javascript:EvalPwdStrength(document.forms[0],this.value);' onblur='javascript:EvalPwdStrength(document.forms[0],this.value);'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100' align='right'><strong>密码强度：</strong></td>" & vbCrLf
    Response.Write "      <td>" & ShowPwdStrength & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100' align='right'><strong>确认密码：</strong></td>" & vbCrLf
    Response.Write "      <td><input type='password' name='PwdConfirm'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Modify'>" & vbCrLf
    Response.Write "        <input  type='submit' name='Submit' value=' 确 定 ' style='cursor:hand;'>" & vbCrLf
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Index_Main.asp'"" style='cursor:hand;'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub ModifyPwd()
    Dim rs, sql
    Dim Password, PwdConfirm
    
    Password = Trim(Request("Password"))
    PwdConfirm = Trim(Request("PwdConfirm"))
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>新密码不能为空！</li>"
    End If
    If PwdConfirm <> Password Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>确认密码必须与新密码相同！</li>"
    End If
    If CheckBadChar(Password) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>新密码中含有非法字符！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    sql = "Select * from PE_Admin where AdminName='" & AdminName & "'"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不存在此管理员！</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs("Password") = MD5(Password, 16)
    rs.Update
    rs.Close
    Set rs = Nothing
    
    Call WriteSuccessMsg("修改密码成功！下次登录时记得换用新密码哦！", ComeUrl)
End Sub

Function ShowPwdStrength()
    Dim strStrength
    strStrength = strStrength & "<script language='JavaScript' src='PwdStrength.js'></script>" & vbCrLf
    strStrength = strStrength & "<script language='JavaScript'>" & vbCrLf
    strStrength = strStrength & "<!--" & vbCrLf
    strStrength = strStrength & "window.onerror = ignoreError;" & vbCrLf
    strStrength = strStrength & "function ignoreError(){return true;}" & vbCrLf
    strStrength = strStrength & "function EvalPwdStrength(oF,sP){" & vbCrLf
    strStrength = strStrength & "  PadPasswd(oF,sP.length*2);" & vbCrLf
    strStrength = strStrength & "  if(ClientSideStrongPassword(sP,gSimilarityMap,gDictionary)){DispPwdStrength(3,'cssStrong');}" & vbCrLf
    strStrength = strStrength & "  else if(ClientSideMediumPassword(sP,gSimilarityMap,gDictionary)){DispPwdStrength(2,'cssMedium');}" & vbCrLf
    strStrength = strStrength & "  else if(ClientSideWeakPassword(sP,gSimilarityMap,gDictionary)){DispPwdStrength(1,'cssWeak');}" & vbCrLf
    strStrength = strStrength & "  else{DispPwdStrength(0,'cssPWD');}" & vbCrLf
    strStrength = strStrength & "}" & vbCrLf
    strStrength = strStrength & "function PadPasswd(oF,lPwd){" & vbCrLf
    strStrength = strStrength & "  if(typeof oF.PwdPad=='object'){var sPad='IfYouAreReadingThisYouHaveTooMuchFreeTime';var lPad=sPad.length-lPwd;oF.PwdPad.value=sPad.substr(0,(lPad<0)?0:lPad);}" & vbCrLf
    strStrength = strStrength & "}" & vbCrLf
    strStrength = strStrength & "function DispPwdStrength(iN,sHL){" & vbCrLf
    strStrength = strStrength & "  if(iN>3){ iN=3;}for(var i=0;i<4;i++){ var sHCR='cssPWD';if(i<=iN){ sHCR=sHL;}if(i>0){ GEId('idSM'+i).className=sHCR;}GEId('idSMT'+i).style.display=((i==iN)?'inline':'none');}" & vbCrLf
    strStrength = strStrength & "}" & vbCrLf
    strStrength = strStrength & "function GEId(sID){return document.getElementById(sID);}" & vbCrLf
    strStrength = strStrength & "//-->" & vbCrLf
    strStrength = strStrength & "</script>" & vbCrLf
    strStrength = strStrength & "<style>" & vbCrLf
    strStrength = strStrength & "input{FONT-FAMILY:宋体;FONT-SIZE: 9pt;}" & vbCrLf
    strStrength = strStrength & ".cssPWD{background-color:#EBEBEB;border-right:solid 1px #BEBEBE;border-bottom:solid 1px #BEBEBE;}" & vbCrLf
    strStrength = strStrength & ".cssWeak{background-color:#FF4545;border-right:solid 1px #BB2B2B;border-bottom:solid 1px #BB2B2B;}" & vbCrLf
    strStrength = strStrength & ".cssMedium{background-color:#FFD35E;border-right:solid 1px #E9AE10;border-bottom:solid 1px #E9AE10;}" & vbCrLf
    strStrength = strStrength & ".cssStrong{background-color:#3ABB1C;border-right:solid 1px #267A12;border-bottom:solid 1px #267A12;}" & vbCrLf
    strStrength = strStrength & ".cssPWT{width:132px;}" & vbCrLf
    strStrength = strStrength & "</style>" & vbCrLf
    strStrength = strStrength & "<table cellpadding='0' cellspacing='0' class='cssPWT' style='height:16px'><tr valign='bottom'><td id='idSM1' width='33%' class='cssPWD' align='center'><span style='font-size:1px'>&nbsp;</span><span id='idSMT1' style='display:none;'>弱</span></td><td id='idSM2' width='34%' class='cssPWD' align='center' style='border-left:solid 1px #fff'><span style='font-size:1px'>&nbsp;</span><span id='idSMT0' style='display:inline;font-weight:normal;color:#666'>无</span><span id='idSMT2' style='display:none;'>中</span></td><td id='idSM3' width='33%' class='cssPWD' align='center' style='border-left:solid 1px #fff'><span style='font-size:1px'>&nbsp;</span><span id='idSMT3' style='display:none;'>强</span></td></tr></table>"
    ShowPwdStrength = strStrength
End Function
%>
