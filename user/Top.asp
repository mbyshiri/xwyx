<!--#include file="../Include/PowerEasy.Edition.asp"-->
<SCRIPT language='JavaScript1.2' src='../js/stm31.js' type='text/javascript'></SCRIPT>

<%
Dim rsTemplateProject, DefaultTemplateProjectName
Set rsTemplateProject = Conn.Execute("Select TemplateProjectName from PE_TemplateProject where IsDefault = " & PE_True & "")
If rsTemplateProject.Bof And rsTemplateProject.Eof Then
    DefaultTemplateProjectName = ""
Else
    DefaultTemplateProjectName = rsTemplateProject(0)
End If
rsTemplateProject.Close
Set rsTemplateProject = Nothing

If DefaultTemplateProjectName = "����2006��������" Then
%>

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
          <td width="83%" height=80><img src="<%=InstallDir&BannerUrl%>" width="468" height="60"></td>
          <td width="17%">
            <table height=89 cellSpacing=0 cellPadding=0 width=94 background=<%=InstallDir%>Skin/images/topr.gif border=0>
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
<table cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <tr>
    <td class=menu_s align=middle><%=GetChannelList(0)%></td>
  </tr>
  <tr>
    <td><IMG height=7 src="../skin/Ocean/menu_bg2.jpg" width=778></td>
  </tr>
  <tr>
    <td background=<%=InstallDir%>Skin/images/addr_line.jpg height=4></td>
  </tr>
</table>
<table class='top_tdbgall' style='word-break: break-all' cellSpacing='0' cellPadding='0' width='760' align='center' border='0'>
  <!--Ƶ����ʾ����-->
  <!--��վLogo��banner��ʾ����-->
  <tr>
    <td><table width='100%' border='0' cellpadding='0' cellspacing='0' background='images/contmenu_bg.gif'>
      <tr>
        <td width='8'><img src='images/contmenu1.gif' width='8' height='45'></td>
        <td width='160' align='right'><img src='images/contmenu.gif' width='151' height='45'></td>
        <td valign='bottom'><table width='100%' height='29' border='0' cellpadding='0' cellspacing='0'>
          <tr>
            <td>

<% ElseIf DefaultTemplateProjectName = "����2006����ϵ��" Then%>

<table class='top_tdbgall' style='word-break: break-all' cellSpacing='0' cellPadding='0' width='760' align='center' border='0'>
  <!--�����վ����-->
  <tr>
    <td class=top_top></td>
  </tr>
  <!--Ƶ����ʾ����-->
  <tr>
    <td><table class='top_Channel' cellSpacing='0' cellPadding='0' width='100%' border='0'>
      <tr>
        <td align=right><%=GetChannelList(0)%></td>
      </tr>
    </table></td>
  </tr>
  <!--��վLogo��banner��ʾ����-->
  <tr>
    <td align=center><a href='#' title='��ҳ' target='_blank'><img src='<%=InstallDir%>Skin/Elegance/logo.gif' width='180' height='60' border='0'></a></td>
  </tr>
  <tr>
    <td><table width='100%' border='0' cellpadding='0' cellspacing='0' background='images/contmenu_bg.gif'>
      <tr>
        <td width='8'><img src='images/contmenu1.gif' width='8' height='45'></td>
        <td width='160' align='right'><img src='images/contmenu.gif' width='151' height='45'></td>
        <td valign='bottom'><table width='100%' height='29' border='0' cellpadding='0' cellspacing='0'>
          <tr>
            <td>
<% ElseIf DefaultTemplateProjectName = "����2006��֮��ģ�巽��" Then%>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
    <tr><td colspan="3" height="6"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td></tr>
    <tr><td width="14" height="34"><img src="<%=InstallDir%>Skin/sealove/Top_01Left.gif"></td>
        <td background="<%=InstallDir%>Skin/sealove/Top_01BG.gif" align="right" style="color:#FFFFFF"><%=GetChannelList(0)%></td>
        <td width="14"><img src="<%=InstallDir%>Skin/sealove/Top_01Right.gif"></td>
    </tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
    <tr><td width="4" background="<%=InstallDir%>Skin/sealove/Top_02Left.gif"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td>
        <td background="<%=InstallDir%>Skin/sealove/Top_02BG.gif">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr><td width="280" height="90" align="center"><a href='<%=SiteUrl%>' title='<%=SiteName%>' target='_blank'><img src='<%=InstallDir%>skin/sealove/PElogo_sealove.gif' border='0'></a></td>
                <td align="center"><a href=\"http://www.powereasy.net\" target=\"_blank\"><img src="<%=InstallDir%>Skin/sealove/PowerEasy_TOP.gif" border=\"0\"></a></td>
            </tr>
        </table>
        </td>
        <td width="4" background="<%=InstallDir%>Skin/sealove/Top_02Right.gif"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td>
    </tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
    <tr><td width="17" background="<%=InstallDir%>Skin/sealove/Top_03Left.gif"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td>
        <td  background="<%=InstallDir%>Skin/sealove/Top_03BG.gif">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr><td width="20" height="33"><img src="<%=InstallDir%>Skin/sealove/icon01.gif"></td>
                <td width="60">���¹��棺</td>
                <td width="400"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td>
                <td align="right"><script language="JavaScript" type="text/JavaScript" src="<%=InstallDir%>js/date.js"></script></td>
            </tr>
        </table>
        </td>
        <td width="17" background="<%=InstallDir%>Skin/sealove/Top_03Right.gif"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td>
    </tr>
    <tr><td colspan="3" height="5"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td></tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
    <tr><td width="15"><img src="<%=InstallDir%>Skin/sealove/Main_TopLeft.gif"></td>
        <td height="11" background="<%=InstallDir%>Skin/sealove/Main_TopBG.gif"><img src="<%=InstallDir%>Skin/sealove/space.gif"></td>
        <td width="15"><img src="<%=InstallDir%>Skin/sealove/Main_TopRight.gif"></td>
    </tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr><td width="8" background="<%=InstallDir%>Skin/sealove/Main_Left.gif"></td>
        <td>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr><td width="135" height="60"><img src="<%=InstallDir%>Skin/sealove/Main_Search.gif" alt="վ������"></td>
                <td><table cellSpacing=0 cellPadding=0 border=0>
                    <FORM name=search action=<%=InstallDir%>search.asp method=post>
                    <tr><td align=middle><Input id=Keyword maxLength=50 value=�ؼ��� name=Keyword></td>
                        <td align="center" width="55"><input name=Submit id=Submit type='image' src='<%=InstallDir%>Skin/sealove/Icon_Search.gif' style='BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px'></td>
                        <td align=middle><Input type=radio CHECKED value=Article name=ModuleName> ����
                        <Input type=radio value=Soft name=ModuleName> ����
                        <Input type=radio value=Photo name=ModuleName> ͼƬ
                        <Input id=Field type=hidden value=Title name=Field></td>
                    </tr>
                    </FORM>
                    </table>
                </td>
                <td width="166" align="right"><img src="<%=InstallDir%>Skin/sealove/Main_girl01.gif"></td>
            </tr>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr><td valign="top">
                <table width="98%" align="center" border="0" cellpadding="0" cellspacing="0" background="<%=InstallDir%>Skin/sealove/Path_BG.gif">
                    <tr><td width="9"><img src="<%=InstallDir%>Skin/sealove/Path_Left.gif"></td>
                        <td width="20"><img src="<%=InstallDir%>Skin/sealove/icon02.gif"></td>
                        <td>�����ڵ�λ�ã�&nbsp;<a class='LinkPath' href='<%=SiteUrl%>'><%=SiteName%></a>&nbsp;>>&nbsp;��Ա����</td>
                        <td width="84"><a href="<%=InstallDir%>Reg/User_Reg.asp" target="_blank"><img src="<%=InstallDir%>Skin/sealove/Button_Reg.gif" alt="��Աע��" border="0"></a></td>
                        <td width="9"><img src="<%=InstallDir%>Skin/sealove/Path_Right.gif"></td>
                    </tr>
                </table>
                </td>
                <td width="92" height="48" align="right" valign="top"><img src="<%=InstallDir%>Skin/sealove/Main_girl02.gif"></td>
            </tr>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr><td height="2" bgcolor="#0099BB"></td></tr>
            <tr><td height="1"></td></tr>
            <tr><td height="1" bgcolor="#0099BB"></td></tr>
            <tr><td height="8"></td></tr>
        </table>
        </td>
        <td width="8" background="<%=InstallDir%>Skin/sealove/Main_Right.gif"></td>
    </tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr><td width="8" background="<%=InstallDir%>Skin/sealove/Main_Left.gif"></td>
        <td>

<table class='top_tdbgall' style='word-break: break-all' cellSpacing='0' cellPadding='0' width='760' align='center' border='0'>
  <tr>
    <td><table width='100%' border='0' cellpadding='0' cellspacing='0' background='images/contmenu_bg.gif'>
      <tr>
        <td width='8'><img src='images/contmenu1.gif' width='8' height='45'></td>
        <td width='160' align='right'><img src='images/contmenu.gif' width='151' height='45'></td>
        <td valign='bottom'><table width='100%' height='29' border='0' cellpadding='0' cellspacing='0'>
          <tr>
            <td>
<%Else%>

<table class='top_tdbgall' style='word-break: break-all' cellSpacing='0' cellPadding='0' width='760' align='center' border='0'>
  <!--�����վ����-->
  <tr>
    <td class=top_top></td>
  </tr>
  <!--Ƶ����ʾ����-->
  <tr>
    <td><table class='top_Channel' cellSpacing='0' cellPadding='0' width='100%' border='0'>
      <tr>
        <td align=right><%=GetChannelList(0)%></td>
      </tr>
    </table></td>
  </tr>
  <!--��վLogo��banner��ʾ����-->
  <tr>
    <td align=center><a href='#' title='��ҳ' target='_blank'><img src='<%= InstallDir %>images/logo.gif' width='180' height='60' border='0'></a></td>
  </tr>
  <tr>
    <td><table width='100%' border='0' cellpadding='0' cellspacing='0' background='images/contmenu_bg.gif'>
      <tr>
        <td width='8'><img src='images/contmenu1.gif' width='8' height='45'></td>
        <td width='160' align='right'><img src='images/contmenu.gif' width='151' height='45'></td>
        <td valign='bottom'><table width='100%' height='29' border='0' cellpadding='0' cellspacing='0'>
          <tr>
            <td>
<%End If%>

<%

Response.Write "<script language='JavaScript1.2' type='text/JavaScript'>" & vbCrLf
Response.Write "stm_bm(['uueoehr',400,'','" & InstallDir & "images/blank.gif',0,'','',0,0,0,0,0,1,0,0]);" & vbCrLf
Response.Write "stm_bp('p0',[0,4,0,0,2,2,0,0,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,0,0,'#000000','transparent','',3,0,0,'#000000']);" & vbCrLf
Response.Write "stm_ai('p0i0',[0,'','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p0i1','p0i0',[0,'��Ա������ҳ','','',-1,-1,0,'Index.asp','_self','Index.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf

Response.Write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p0i3','p0i0',[0,'��Ϣ����','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_bp('p1',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf

Dim sqlChannel, rsChannel
sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False
Select Case SystemEdition
Case "CMS", "eShop"
    sqlChannel = sqlChannel & " and ModuleType<4"
Case "GPS", "EPS", "ECS"
    sqlChannel = sqlChannel & " and (ModuleType<4 or ModuleType=8)"
Case "IPS"
    sqlChannel = sqlChannel & " and (ModuleType<4 or ModuleType=6 or ModuleType=7)"
Case "All"
    sqlChannel = sqlChannel & " and (ModuleType<4 or ModuleType>5)"
End Select
sqlChannel = sqlChannel & " order by OrderID"
Set rsChannel = Conn.Execute(sqlChannel)
Do While Not rsChannel.EOF
    ChannelID = rsChannel("ChannelID")
    ChannelName = Trim(rsChannel("ChannelName"))
    ChannelShortName = Trim(rsChannel("ChannelShortName"))
    ChannelDir = Trim(rsChannel("ChannelDir"))
    Select Case rsChannel("ModuleType")
    Case 1
        ModuleName = "Article"
    Case 2
        ModuleName = "Soft"
    Case 3
        ModuleName = "Photo"
    Case 6
        ModuleName = "Supply"
    Case 7
        ModuleName = "House"
    Case 8
        ModuleName = "Job"
    End Select
    If ChannelID = 998 Then
        Response.Write "stm_aix('p1i0','p0i0',[0,'" & ChannelName & "����','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_bpx('p2','p1',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
        Dim rsHouseClass
		Set rsHouseClass = Conn.Execute("select * from PE_HouseConfig")
        Do While Not rsHouseClass.EOF
            Response.Write "stm_aix('p2i0','p1i0',[0,'����" & rsHouseClass("ClassName") & "��Ϣ','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ClassID=" & rsHouseClass("ClassID") & "&Action=Add','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
            Response.Write "stm_aix('p2i0','p1i0',[0,'����" & rsHouseClass("ClassName") & "��Ϣ','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ClassID=" & rsHouseClass("ClassID") & "&Action=Manage','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
            rsHouseClass.MoveNext
        Loop
        Response.Write "stm_ep();" & vbCrLf
    End If

    If ChannelID = 997 Then
        Response.Write "stm_aix('p1i0','p0i0',[0,'�ҵļ�������','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_bpx('p2','p1',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
        Response.Write "stm_aix('p2i0','p1i0',[0,'��ѯְλ��Ϣ','','',-1,-1,0,'../Job/Searchresult.asp','_self','../Job/Searchresult.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_aix('p2i0','p1i0',[0,'ά���ҵļ���','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_aix('p2i0','p1i0',[0,'�������ְλ','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Supply','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Supply' ,'','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_ep();" & vbCrLf
    End If

    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = False And ChannelID <> 997 And ChannelID <> 998 Then '���Ӳ���ʾ����������
        Response.Write "stm_aix('p1i0','p0i0',[0,'" & ChannelName & "����','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_bpx('p2','p1',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
        If CheckUser_ChannelInput() = True Then
            Response.Write "stm_aix('p2i0','p1i0',[0,'���" & ChannelShortName & "','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        End If
        Response.Write "stm_aix('p2i0','p1i0',[0,'����ӵ�" & ChannelShortName & "','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_aix('p2i0','p1i0',[0,'���ղص�" & ChannelShortName & "','','',-1,-1,0,'User_Favorite.asp?ChannelID=" & ChannelID & "','_self','User_Favorite.asp?ChannelID=" & ChannelID & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_aix('p2i0','p1i0',[0,'�����۵�" & ChannelShortName & "','','',-1,-1,0,'User_Comment.asp?ChannelID=" & ChannelID & "','_self','User_Comment.asp?ChannelID=" & ChannelID & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        If rsChannel("ModuleType") = 1 Then
            Response.Write "stm_aix('p2i0','p1i0',[0,'ǩ�����¹���','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Receive&Passed=All','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Receive&Passed=All','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        End If
        Response.Write "stm_ep();" & vbCrLf
    Else
    End If
    rsChannel.MoveNext
Loop
rsChannel.Close
Set rsChannel = Nothing
If FoundInArr(AllModules, "Classroom", ",") Then
    Response.Write "stm_aix('p1i0','p0i0',[0,'�ҳ�ʹ�õǼ�','','',-1,-1,0,'User_Enrol.asp','_self','User_Enrol.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
End If

    
Response.Write "stm_ep();" & vbCrLf

Dim rsChannel_Shop, NoShow_Shop
Set rsChannel_Shop = Conn.Execute("select Disabled from PE_Channel where ModuleType=5")
If Not (rsChannel_Shop.bof And rsChannel_Shop.EOF) Then
    NoShow_Shop = rsChannel_Shop(0)
Else
    NoShow_Shop = True
End If

If NoShow_Shop = False Then
    Response.Write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p0i4','p0i0',[0,'�̳ǹ���','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
    If PE_Clng(UserSetting(30)) = 1 Then
        Response.Write "stm_aix('p2i0','p1i0',[0,'������Ʒ','','',-1,-1,0,'User_Wholesale.asp','_self','User_Wholesale.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    End If
    If GroupType = 4 Then
        Response.Write "stm_aix('p2i0','p0i0',[0,'�Ҵ���Ķ���','','',-1,-1,0,'User_Order.asp?OrderType=1','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵĶ��˵�','','',-1,-1,0,'User_Bill.asp','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Response.Write "stm_aix('p2i0','p0i0',[0,'��Ͷ�߼�¼','','',-1,-1,0,'User_Complain.asp','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    End If
    Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵĶ���','','',-1,-1,0,'User_Order.asp','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵĹ��ﳵ','','',-1,-1,0,'../Shop/ShoppingCart.asp','_blank','../Shop/ShoppingCart.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'���ղص���Ʒ','','',-1,-1,0,'User_Favorite.asp?ChannelID=1000','_self','User_Favorite.asp?ChannelID=1000','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'�����۵���Ʒ','','',-1,-1,0,'User_Comment.asp?ChannelID=1000','_self','User_Comment.asp?ChannelID=1000','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'����֧��','','',-1,-1,0,'../PayOnline/PayOnline.asp','_blank','../PayOnline/PayOnline.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'����֧����ѯ','','',-1,-1,0,'User_Payment.asp','_self','User_Payment.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'�ʽ���ϸ��ѯ','','',-1,-1,0,'User_Bankroll.asp','_self','User_Bankroll.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'���ع�������','','',-1,-1,0,'User_Down.asp','_self','User_Down.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p2i0','p0i0',[0,'��ȡ�����ֵ��','','',-1,-1,0,'User_Exchange.asp?Action=GetCard','_self','User_Exchange.asp?Action=GetCard','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
End If

Response.Write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p0i5','p0i0',[0,'����Ϣ����','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Inbox','_self','User_Message.asp?Action=Manage&ManageType=Inbox','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'׫д����Ϣ','','',-1,-1,0,'User_Message.asp?Action=New','_self','User_Message.asp?Action=New','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'�ռ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Inbox','_self','User_Message.asp?Action=Manage&ManageType=Inbox','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'�ݸ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Outbox','_self','User_Message.asp?Action=Manage&ManageType=Outbox','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'�ѷ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=IsSend','_self','User_Message.asp?Action=Manage&ManageType=IsSend','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'�ϼ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Recycle','_self','User_Message.asp?Action=Manage&ManageType=Recycle','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_ep();" & vbCrLf
Response.Write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p0i7','p0i0',[0,'��ֵ����','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
If UserSetting(18) = 1 Then
    Response.Write "stm_aix('p2i0','p0i0',[0,'�һ�" & PointName & "','','',-1,-1,0,'User_Exchange.asp?Action=Exchange','_self','User_Exchange.asp?Action=Exchange','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
End If
If UserSetting(19) = 1 Then
    Response.Write "stm_aix('p2i0','p0i0',[0,'�һ���Ч��','','',-1,-1,0,'User_Exchange.asp?Action=Valid','_self','User_Exchange.asp?Action=Valid','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
End If
Response.Write "stm_aix('p2i0','p0i0',[0,'��ֵ����ֵ','','',-1,-1,0,'User_Exchange.asp?Action=Recharge','_self','User_Exchange.asp?Action=Recharge','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
If UserSetting(20) = 1 Then
    Response.Write "stm_aix('p2i0','p0i0',[0,'����" & PointName & "','','',-1,-1,0,'User_Exchange.asp?Action=SendPoint','_self','User_Exchange.asp?Action=SendPoint','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
End If
Response.Write "stm_aix('p2i0','p0i0',[0,'����֧����ѯ','','',-1,-1,0,'User_Payment.asp','_self','User_Payment.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'�ʽ���ϸ��ѯ','','',-1,-1,0,'User_Bankroll.asp','_self','User_Bankroll.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'" & PointName & "��ϸ��ѯ','','',-1,-1,0,'User_ConsumeLog.asp','_self','User_ConsumeLog.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'��Ч����ϸ��ѯ','','',-1,-1,0,'User_RechargeLog.asp','_self','User_RechargeLog.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_ep();" & vbCrLf

If UserSetting(25) = 1 Then
    Response.Write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_aix('p0i8','p0i0',[0,'�ҵľۺ�','','',-1,-1,0,'User_Space.asp?Action=Manage','_self','User_Space.asp?Action=Manage','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Response.Write "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
    Dim rsspace, rsitem
    Set rsspace = Conn.Execute("select top 1 Passed from PE_Space where Type=1 and UserID=" & UserID)
    If rsspace.bof And rsspace.EOF Then
        Response.Write "stm_aix('p2i0','p0i0',[0,'����ۺϿռ�','','',-1,-1,0,'User_Space.asp?Action=Add','_self','User_Space.asp?Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
    Else
        If rsspace("Passed") = True Then
            Response.Write "stm_aix('p2i0','p0i0',[0,'��������Ŀ','','',-1,-1,0,'User_Space.asp?Action=Add','_self','User_Space.asp?Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
            Set rsitem = Conn.Execute("select ID,Name,Type from PE_Space where (Type>=3 and Type<=7) and Passed=" & PE_True & " and UserID=" & UserID & " order by Type desc")
            Do While Not rsitem.EOF
                Select Case rsitem("Type")
                Case 3
                    Response.Write "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'д��־','','',-1,-1,0,'User_SpaceDiary.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceDiary.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵ���־����','','',-1,-1,0,'User_SpaceDiary.asp?ID=" & rsitem("ID") & "','_self','User_SpaceDiary.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_ep();" & vbCrLf
                Case 4
                    Response.Write "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�������','','',-1,-1,0,'User_SpaceMusic.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceMusic.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵ����ֹ���','','',-1,-1,0,'User_SpaceMusic.asp?ID=" & rsitem("ID") & "','_self','User_SpaceMusic.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_ep();" & vbCrLf
                Case 5
                    Response.Write "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�������','','',-1,-1,0,'User_SpaceBook.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceBook.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵ�ͼ�����','','',-1,-1,0,'User_SpaceBook.asp?ID=" & rsitem("ID") & "','_self','User_SpaceBook.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_ep();" & vbCrLf
                Case 6
                    Response.Write "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�����ͼƬ','','',-1,-1,0,'User_SpacePhoto.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpacePhoto.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵ�ͼƬ����','','',-1,-1,0,'User_SpacePhoto.asp?ID=" & rsitem("ID") & "','_self','User_SpacePhoto.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_ep();" & vbCrLf
                Case 7
                    Response.Write "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'���������','','',-1,-1,0,'User_SpaceLink.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceLink.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_aix('p2i0','p0i0',[0,'�ҵ����ӹ���','','',-1,-1,0,'User_SpaceLink.asp?ID=" & rsitem("ID") & "','_self','User_SpaceLink.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
                    Response.Write "stm_ep();" & vbCrLf
                End Select
                rsitem.MoveNext
            Loop
            Set rsitem = Nothing
            If UserSetting(28) = 1 Then
            Response.Write "stm_aix('p2i0','p0i0',[0,'�����ռ��ʽ','','',-1,-1,0,'User_Space.asp?Action=Template','_self','User_Space.asp?Action=Template','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
            End If
            Response.Write "stm_aix('p2i0','p0i0',[0,'�鿴�ҵľۺ�','','',-1,-1,0,'../Space/" & UserName & UserID & "/','_blank','../Space/" & UserName & UserID & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        Else
            Response.Write "stm_aix('p2i0','p0i0',[0,'�ۺϿռ������...','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
        End If
    End If
End If


Response.Write "stm_ep();" & vbCrLf
Response.Write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p0i6','p0i0',[0,'�û�����','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
Response.Write "stm_aix('p0i0','p0i0',[0,'�����б�','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'��Ա�б�','','',-1,-1,0,'User_Friend.asp','_self','User_Friend.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'��ӳ�Ա','','',-1,-1,0,'User_Friend.asp?Action=AddFriend','_self','User_Friend.asp?Action=AddFriend','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'��������','','',-1,-1,0,'User_Friend.asp?Action=CreateNewGroup','_self','User_Friend.asp?Action=CreateNewGroup','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p2i0','p0i0',[0,'�������','','',-1,-1,0,'User_Friend.asp?Action=ManageGroup','_self','User_Friend.asp?Action=ManageGroup','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_ep();" & vbCrLf
Response.Write "stm_aix('p0i0','p0i0',[0,'�޸�����','','',-1,-1,0,'User_Info.asp?Action=ModifyPwd','_self','User_Info.asp?Action=ModifyPwd','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p0i0','p0i0',[0,'�޸���Ϣ','','',-1,-1,0,'User_Info.asp?Action=Modify','_self','User_Info.asp?Action=Modify','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
If UserType = 0 Then
    Response.Write "stm_aix('p0i0','p0i0',[0,'ע���ҵ���ҵ','','',-1,-1,0,'User_Info.asp?Action=RegCompany','_self','User_Info.asp?Action=RegCompany','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
End If
Response.Write "stm_aix('p0i0','p0i0',[0,'�ʼ����Ĺ���','','',-1,-1,0,'User_mailreg.asp','_self','User_mailreg.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_aix('p0i0','p0i0',[0,'�˳���¼','','',-1,-1,0,'User_Logout.asp','_self','User_Logout.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
Response.Write "stm_ep();" & vbCrLf
Response.Write "stm_em();" & vbCrLf
Response.Write "</script>" & vbCrLf
%>

            </td>
          </tr>
          <tr>
            <td height='4'></td>
          </tr>
        </table></td>
        <td width='6'><img src='images/contmenu2.gif' width='6' height='45'></td>
      </tr>
    </table></td>
  </tr>
</table>



<%
Dim tMessageID, rsMessage
If request("Action") <> "ReadInbox" Then
    Set rsMessage = Conn.Execute("select Min(Id) from PE_Message where incept='" & UserName & "'and delR=0 and flag=0 and IsSend=1")
    If IsNull(rsMessage(0)) Then
        tMessageID = 0
    Else
        tMessageID = rsMessage(0)
    End If
    Set rsMessage = Nothing
    If tMessageID > 0 Then
        Response.Write "<script LANGUAGE='JavaScript'>" & vbCrLf
        Response.Write "var url = 'User_ReadMessage.asp?MessageID=" & tMessageID & "';" & vbCrLf
        Response.Write "window.open (url, 'newmessage', 'height=440, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')" & vbCrLf
        Response.Write "</script>" & vbCrLf
    End If
End If
%>