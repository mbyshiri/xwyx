<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Edition.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 1      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��


Response.Write "<html><head><title>��վ����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<script language='JavaScript'>" & vbCrLf
Response.Write "function SelectColor(sEL,form){" & vbCrLf
Response.Write "    var dEL = document.all(sEL);" & vbCrLf
Response.Write "    var url = '../Editor/editor_selcolor.asp?color='+encodeURIComponent(sEL);" & vbCrLf
Response.Write "    var arr = showModalDialog(url,window,'dialogWidth:280px;dialogHeight:250px;help:no;scroll:no;status:no');" & vbCrLf
Response.Write "    if (arr) {" & vbCrLf
Response.Write "        form.value=arr;" & vbCrLf
Response.Write "        sEL.style.backgroundColor=arr;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
Call ShowPageTitle("�� վ �� Ϣ �� ��", 10001)
Response.Write "</table>" & vbCrLf

If Action = "SaveConfig" Then
    Call SaveConfig
    Call WriteEntry(1, AdminName, "�޸���վ��Ϣ����")
Else
    Call ModifyConfig
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn




Sub ModifyConfig()
    Dim sqlConfig, rsConfig
    
    sqlConfig = "select * from PE_Config"
    Set rsConfig = Server.CreateObject("ADODB.Recordset")
    rsConfig.Open sqlConfig, Conn, 1, 3
    If rsConfig.BOF And rsConfig.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��վ�������ݶ�ʧ����ʹ�ó�ʼ���ݿ���лָ���</li>"
        rsConfig.Close
        Set rsConfig = Nothing
        Exit Sub
    End If
    
    Dim RegFields_MustFill, Modules
    RegFields_MustFill = rsConfig("RegFields_MustFill")
    Modules = rsConfig("Modules")

    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function IsDigit()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  return ((event.keyCode >= 48) && (event.keyCode <= 57));" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form name='myform' id='myform' method='POST' action='Admin_SiteConfig.asp' >" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'>"
    Response.Write "<td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��վ��Ϣ</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��վѡ��</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(2)'>��Աѡ��</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(3)'>�ʼ�ѡ��</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(4)'>����ͼѡ��</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(5)'>����ѡ��</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(6)'>�̳�ѡ��</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(7)'>���Ա�ѡ��</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(8)'>Rss/WAP����</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(9)'>�ֻ���������</td>" & vbCrLf
    Response.Write "<td>&nbsp;</td></tr></table>"

    
    Response.Write "<table width='100%' border='0' cellpadding='5' cellspacing='1' Class='border'><tr><td class='tdbg'>" & vbCrLf
    Response.Write "<table width='95%' border='0' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ���ƣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteName' type='text' id='SiteName' value='" & rsConfig("SiteName") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ���⣺</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteTitle' type='text' id='SiteTitle' value='" & rsConfig("SiteTitle") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ��ַ��</strong><br>����д����URL��ַ</td>" & vbCrLf
    Response.Write "      <td><input name='SiteUrl' type='text' id='SiteUrl' value='" & rsConfig("SiteUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><font color=red><strong>��װĿ¼��</strong><br>ϵͳ��װĿ¼������ڸ�Ŀ¼��λ�ã�<br>ϵͳ���Զ������ȷ��·��������Ҫ�ֹ��������á�</font></td>" & vbCrLf
    Response.Write "      <td><input name='InstallDir' type='text' id='InstallDir' value='" & InstallDir & "' size='40' maxlength='30' readonly></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>LOGO��ַ��</strong><br>����д����URL��ַ</td>" & vbCrLf
    Response.Write "      <td><input name='LogoUrl' type='text' id='LogoUrl' value='" & rsConfig("LogoUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>Banner��ַ��</strong><br>����д����URL��ַ</td>" & vbCrLf
    Response.Write "      <td><input name='BannerUrl' type='text' id='BannerUrl' value='" & rsConfig("BannerUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>վ��������</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterName' type='text' id='WebmasterName' value='" & rsConfig("WebmasterName") & "' size='40' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>վ�����䣺</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterEmail' type='text' id='WebmasterEmail' value='" & rsConfig("WebmasterEmail") & "' size='40' maxlength='100'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ȩ��Ϣ��</strong><br>֧��HTML��ǣ�����ʹ��˫����</td>" & vbCrLf
    Response.Write "      <td><textarea name='Copyright' cols='60' rows='4' id='Copyright'>" & rsConfig("Copyright") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վMETA�ؼ��ʣ�</strong><br>��������������õĹؼ���<br>����ؼ�������,�ŷָ�</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'>" & rsConfig("Meta_Keywords") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վMETA��ҳ������</strong><br>��������������õ���ҳ����<br>�����������,�ŷָ�</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'>" & rsConfig("Meta_Description") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ���ʾ��վƵ����</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowSiteChannel' value='1' " & IsRadioChecked(rsConfig("ShowSiteChannel"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowSiteChannel' value='0' " & IsRadioChecked(rsConfig("ShowSiteChannel"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ���ʾ�����¼���ӣ�</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowAdminLogin' value='1' " & IsRadioChecked(rsConfig("ShowAdminLogin"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowAdminLogin' value='0' " & IsRadioChecked(rsConfig("ShowAdminLogin"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ񱣴�Զ��ͼƬ�����أ�</strong><br>�����������վ�ϸ��Ƶ������а���ͼƬ����ͼƬ���Ƶ���վ��������</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableSaveRemote' value='1' " & IsRadioChecked(rsConfig("EnableSaveRemote"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableSaveRemote' value='0' " & IsRadioChecked(rsConfig("EnableSaveRemote"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ񿪷������������룺</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableLinkReg' value='1' " & IsRadioChecked(rsConfig("EnableLinkReg"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableLinkReg' value='0' " & IsRadioChecked(rsConfig("EnableLinkReg"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ�ͳ���������ӵ������</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableCountFriendSiteHits' value='1' " & IsRadioChecked(rsConfig("EnableCountFriendSiteHits"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableCountFriendSiteHits' value='0' " & IsRadioChecked(rsConfig("EnableCountFriendSiteHits"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ�ʹ��������������룺</strong><br>��ѡ���ǣ����Ա��¼��̨ʱʹ��������������룬�ʺ����ɵȳ�������ʹ�á�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableSoftKey' value='1' " & IsRadioChecked(rsConfig("EnableSoftKey"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableSoftKey' value='0' " & IsRadioChecked(rsConfig("EnableSoftKey"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ�ʹ��Ƶ������Ŀ��ר���������ݣ�</strong><br>��ѡ���ǣ�Ƶ������Ŀ��ר������������������ѡ�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='IsCustom_Content' value='1' " & IsRadioChecked(rsConfig("IsCustom_Content"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='IsCustom_Content' value='0' " & IsRadioChecked(rsConfig("IsCustom_Content"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
	
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ�������վ����Ͷ�幦�ܣ�</strong><br>��ѡ���ǣ���վ����������Ͷ���û��飬����Ͷ��ģ�壬ǰ̨����Ͷ�幦�ܡ�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowAnonymous' value='1' " & IsRadioChecked(PE_CBool(rsConfig("ShowAnonymous")), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowAnonymous' value='0' " & IsRadioChecked(PE_CBool(rsConfig("ShowAnonymous")), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf	
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>FSO(FileSystemObject)��������ƣ�</strong><br>ĳЩ��վΪ�˰�ȫ����FSO��������ƽ��и����Դﵽ����FSO��Ŀ�ġ���������վ���������ģ����ڴ�������Ĺ������ơ�</td>" & vbCrLf
    Response.Write "      <td><input name='objName_FSO' type='text' id='objName_FSO' value='" & rsConfig("objName_FSO") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��̨����Ŀ¼��</strong><br>Ϊ�˰�ȫ���������޸ĺ�̨����Ŀ¼��Ĭ��ΪAdmin�����Ĺ��Ժ���Ҫ�����ô˴�</td>" & vbCrLf
    Response.Write "      <td><input name='AdminDir' type='text' id='AdminDir' value='" & rsConfig("AdminDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ���Ŀ¼��</strong><br>Ϊ�˲��ù���������������վ�Ĺ�棬�������޸Ĺ��JS�Ĵ��Ŀ¼��Ĭ��ΪAD�����Ĺ��Ժ���Ҫ�����ô˴�</td>" & vbCrLf
    Response.Write "      <td><input name='ADDir' type='text' id='ADDir' value='" & rsConfig("ADDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�������洰�ڵļ��ʱ�䣺</strong><br>��СʱΪ��λ��Ϊ0ʱÿ��ˢ��ҳ��ʱ���������档</td>" & vbCrLf
    Response.Write "      <td><input name='AnnounceCookieTime' type='text' id='AnnounceCookieTime' value='" & rsConfig("AnnounceCookieTime") & "' size='10' maxlength='10'> Сʱ</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ�ȵ�ĵ������Сֵ��</strong><br>ֻ�е�����ﵽ����ֵ���Ż���Ϊ��վ���ȵ�������ʾ��</td>" & vbCrLf
    Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='" & rsConfig("HitsOfHot") & "' size='10' maxlength='10'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ģ�����ѡ�</strong><br>������վ���õ�ģ�顣</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <table width='100%'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Advertisement'" & IsModulesSelected(Modules, "Advertisement") & ">��վ������</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='FriendSite'" & IsModulesSelected(Modules, "FriendSite") & ">�������ӹ���</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Announce'" & IsModulesSelected(Modules, "Announce") & ">��վ�������</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Vote'" & IsModulesSelected(Modules, "Vote") & ">��վ�������</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf

    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='KeyLink'" & IsModulesSelected(Modules, "KeyLink") & ">վ�����ӹ���</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Rtext'" & IsModulesSelected(Modules, "Rtext") & ">�ַ����˹���</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Collection'" & IsModulesSelected(Modules, "Collection") & ">�ɼ�����</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='SMS'" & IsModulesSelected(Modules, "SMS") & ">�ֻ�����</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    If SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "IPS" Or SystemEdition = "All" Then
    Response.Write "          <td><input name='Modules' type='checkbox' value='Survey'" & IsModulesSelected(Modules, "Survey") & ">�ʾ�������</td>" & vbCrLf
    End If
    If SystemEdition = "IPS" Or SystemEdition = "All" Then
    Response.Write "            <td><input name='Modules' type='checkbox' value='Supply'" & IsModulesSelected(Modules, "Supply") & ">������Ϣ����</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='House'" & IsModulesSelected(Modules, "House") & ">�������Ĺ���</td>" & vbCrLf
    End If
    If SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "All" Then
    Response.Write "            <td><input name='Modules' type='checkbox' value='Job'" & IsModulesSelected(Modules, "Job") & ">�˲���Ƹ����</td>" & vbCrLf
    End If
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    If SystemEdition = "eShop" Or SystemEdition = "ECS" Or SystemEdition = "All" Then
    Response.Write "            <td><input name='Modules' type='checkbox' value='CRM'" & IsModulesSelected(Modules, "CRM") & ">�ͻ���ϵ����</td>" & vbCrLf
    End If
    If SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "All" Then
    Response.Write "          <td><input name='Modules' type='checkbox' value='Classroom'" & IsModulesSelected(Modules, "Classroom") & ">�ҳ��Ǽǹ���</td>" & vbCrLf
    End If
    If SystemEdition = "EPS" Or SystemEdition = "All" Then
    Response.Write "          <td><input name='Modules' type='checkbox' value='Sdms'" & IsModulesSelected(Modules, "Sdms") & ">ѧ��ѧ������</td>" & vbCrLf
    End If
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><font color=red>��վ��ҳ����չ����</font></strong><br>��ѡ��ǰ�������������վ��ҳ������HTML���ܡ�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='0' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 0) & ">.html &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='1' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 1) & ">.htm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='2' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 2) & ">.shtml &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='3' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 3) & ">.shtm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='4' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 4) & ">.asp " & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><font color=red>ȫվר�����չ����</font></strong><br>��ѡ��ǰ�����������ȫվר�������HTML���ܡ�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='0' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 0) & ">.html &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='1' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 1) & ">.htm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='2' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 2) & ">.shtml &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='3' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 3) & ">.shtm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='4' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 4) & ">.asp " & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>���ӵ�ַ��ʽ��</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='SiteUrlType' type='radio' value='0' " & IsRadioChecked(rsConfig("SiteUrlType"), 0) & "> ���·�������磺&lt;a href='/News/200509/1358.html'&gt;����&lt;/a&gt;��<br>&nbsp;&nbsp;&nbsp;&nbsp;��һ����վ�ж������ʱ��һ����ô˷�ʽ<br>&nbsp;&nbsp;&nbsp;&nbsp;��һ����վ�ж��������վʱ��������ô˷�ʽ<br>" & vbCrLf
    Response.Write "        <input name='SiteUrlType' type='radio' value='1' " & IsRadioChecked(rsConfig("SiteUrlType"), 1) & "> ����·�������磺&lt;a href='http://www.powereasy.net/News/200509/1358.html'&gt;����&lt;/a&gt;��<br>&nbsp;&nbsp;&nbsp;&nbsp;��Ҫ��Ƶ����Ϊ��վ��������ʱ������ʹ�ô˷�ʽ<br>&nbsp;&nbsp;&nbsp;&nbsp;Ҫʹ�ô˷�ʽ���������վURL������ȷ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�����޶���ʽ��</strong><br><font color='red'>�˹���ֻ��ASP���ʷ�ʽ��Ч���������ǰ������HTML�ļ��������ô˹��ܺ���ЩHTML�ļ��Կ��Է��ʣ������ֹ�ɾ����������ʹ�ô˹������Ƶ������Ŀ�������µ�Ȩ�����ú�����HTML��ʽ���ﵽ��վ�޶�IP���ʣ�����ֻ����Ȩ�����õ����ݽ���IP�޶���</font></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='0' " & IsRadioChecked(rsConfig("LockIPType"), 0) & ">  �����������޶����ܣ��κ�IP�����Է��ʱ�վ��<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='1' " & IsRadioChecked(rsConfig("LockIPType"), 1) & ">  �������ð�������ֻ����������е�IP���ʱ�վ��<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='2' " & IsRadioChecked(rsConfig("LockIPType"), 2) & ">  �������ú�������ֻ��ֹ�������е�IP���ʱ�վ��<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='3' " & IsRadioChecked(rsConfig("LockIPType"), 3) & ">  ͬʱ���ð�����������������ж�IP�Ƿ��ڰ������У�������ڣ����ֹ���ʣ�����������ж��Ƿ��ں������У����IP�ں����������ֹ���ʣ�����������ʡ�<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='4' " & IsRadioChecked(rsConfig("LockIPType"), 4) & ">  ͬʱ���ð�����������������ж�IP�Ƿ��ں������У�������ڣ���������ʣ�����������ж��Ƿ��ڰ������У����IP�ڰ���������������ʣ������ֹ���ʡ� " & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "   <tr class='tdbg'>     " & vbCrLf
    Response.Write "     <td width='40%' class='tdbg5'>               <strong>IP�ΰ�����</strong>��<br>" & vbCrLf
    Response.Write "      (ע����Ӷ���޶�IP�Σ�����<font color='red'>�س�</font>�ָ��� <br>" & vbCrLf
    Response.Write "      ����IP�ε���д��ʽ���м�����Ӣ���ĸ�С������ӣ��� " & vbCrLf
    Response.Write "      <font color='red'>219.100.93.32----219.100.93.255</font> ���޶���IP 219.100.93.32 ��IP 219.100.93.255 ���IP�εķ��ʡ���ҳ��Ϊasp��ʽʱ����Ч��) </td>      " & vbCrLf
    Response.Write "     <td class='tdbg'>" & vbCrLf

    Response.Write " <textarea name='LockIPWhite' cols='60' rows='8' id='LockIP'>" & vbCrLf
    Dim rsLockIP, arrLockIP, i, arrLockIPCut
    If InStr(rsConfig("LockIP"), "|||") > 0 Then
        rsLockIP = Split(rsConfig("LockIP"), "|||")
        If InStr(rsLockIP(0), "$$$") > 0 Then
            arrLockIP = Split(Trim(rsLockIP(0)), "$$$")
            For i = 0 To UBound(arrLockIP)
                arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
                Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
                If i < UBound(arrLockIP) Then Response.Write Chr(10)
            Next
        ElseIf rsLockIP(0) <> "" Then
            arrLockIPCut = Split(Trim(rsLockIP(0)), "----")
            Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
        End If
    End If
    Response.Write "</textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "   <tr class='tdbg'>     " & vbCrLf
    Response.Write "     <td width='40%' class='tdbg5'>               <strong>IP�κ�����</strong>��<br>" & vbCrLf
    Response.Write "      (ע��ͬ�ϡ�) <br>" & vbCrLf
    Response.Write "      </td>      " & vbCrLf
    Response.Write "     <td class='tdbg'>" & vbCrLf

    Response.Write " <textarea name='LockIPBlack' cols='60' rows='8' id='LockIP'>" & vbCrLf

    If InStr(rsConfig("LockIP"), "|||") > 0 Then
        rsLockIP = Split(rsConfig("LockIP"), "|||")
        If InStr(rsLockIP(1), "$$$") > 0 Then
            arrLockIP = Split(Trim(rsLockIP(1)), "$$$")
            For i = 0 To UBound(arrLockIP)
                arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
                Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
                If i < UBound(arrLockIP) Then Response.Write Chr(10)
            Next
        ElseIf rsLockIP(1) <> "" Then
            arrLockIPCut = Split(Trim(rsLockIP(1)), "----")
            Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
        End If
    End If
    Response.Write "</textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ������»�Աע�᣺</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableUserReg' value='1' " & IsRadioChecked(rsConfig("EnableUserReg"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableUserReg' value='0' " & IsRadioChecked(rsConfig("EnableUserReg"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע���Ƿ���Ҫ�ʼ���֤��</strong><br>��ѡ���ǡ������Աע���ϵͳ�ᷢһ�������֤����ʼ����˻�Ա����Ա������ͨ���ʼ���֤�����������Ϊ��ʽע���Ա</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EmailCheckReg' value='1' " & IsRadioChecked(rsConfig("EmailCheckReg"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EmailCheckReg' value='0' " & IsRadioChecked(rsConfig("EmailCheckReg"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע���Ƿ���Ҫ����Ա��֤��</strong><br>��ѡ���ǣ����Ա������ͨ������Ա��֤����������ɹ���ʽע���Ա��</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='AdminCheckReg' value='1' " & IsRadioChecked(rsConfig("AdminCheckReg"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='AdminCheckReg' value='0' " & IsRadioChecked(rsConfig("AdminCheckReg"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ÿ��Email�Ƿ�����ע���Σ�</strong><br>��ѡ���ǣ�������ͬһ��Email����ע������Ա��</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableMultiRegPerEmail' value='1' " & IsRadioChecked(rsConfig("EnableMultiRegPerEmail"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableMultiRegPerEmail' value='0' " & IsRadioChecked(rsConfig("EnableMultiRegPerEmail"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Աע��ʱ�Ƿ�������֤�빦�ܣ�</strong><br>������֤�빦�ܿ�����һ���̶��Ϸ�ֹ����Ӫ�������ע����Զ�ע��</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableCheckCodeOfReg' value='1' " & IsRadioChecked(rsConfig("EnableCheckCodeOfReg"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableCheckCodeOfReg' value='0' " & IsRadioChecked(rsConfig("EnableCheckCodeOfReg"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Աע��ʱ�Ƿ����ûش�������֤���ܣ�</strong><br>���ô˹��ܣ��������̶��Ϸ�ֹ����Ӫ�������ע����Զ�ע�ᣬҲ��������ĳЩ���ⳡ�ϣ���ֹ�޹���Աע���Ա��</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableQAofReg' value='1' " & IsRadioChecked(rsConfig("EnableQAofReg"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableQAofReg' value='0' " & IsRadioChecked(rsConfig("EnableQAofReg"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ����û�Ա����ģ�幦�ܣ�</strong><br>������û�Ա����ģ�壬�����ڻ�Աģ��������޸���ҳģ�壬����Լ��޸Ĺ���Ա����ģ�壬������Ӻ���Ӧģ�幦��֮�������ô˹��ܡ�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowUserModel' value='1' " & IsRadioChecked(PE_CBool(rsConfig("ShowUserModel")), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowUserModel' value='0' " & IsRadioChecked(PE_CBool(rsConfig("ShowUserModel")), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf	
    Dim arrQA

    arrQA = Split(rsConfig("QAofReg") & "", "$$$")
    If UBound(arrQA) <> 5 Then arrQA = Split("����һ$$$��һ$$$�����$$$�𰸶�$$$������$$$����", "$$$")
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����һ��</strong><br>���������֤���ܣ�������һ�ʹ𰸱�����д��</td>" & vbCrLf
    Response.Write "      <td>���⣺<input type='text' name='RegQuestion1' value='" & Trim(arrQA(0)) & "' size='50'><br>�𰸣�<input type='text' name='RegAnswer1' value='" & Trim(arrQA(1)) & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�������</strong><br>���������ѡ��</td>" & vbCrLf
    Response.Write "      <td>���⣺<input type='text' name='RegQuestion2' value='" & Trim(arrQA(2)) & "' size='50'><br>�𰸣�<input type='text' name='RegAnswer2' value='" & Trim(arrQA(3)) & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��������</strong><br>����������ѡ��</td>" & vbCrLf
    Response.Write "      <td>���⣺<input type='text' name='RegQuestion3' value='" & Trim(arrQA(4)) & "' size='50'><br>�𰸣�<input type='text' name='RegAnswer3' value='" & Trim(arrQA(5)) & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע��ʱ�û��������ַ�����</strong></td>" & vbCrLf
    Response.Write "      <td><input name='UserNameLimit' type='text' id='UserNameLimit' value='" & rsConfig("UserNameLimit") & "' size='6' maxlength='5'> ���ַ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע��ʱ�û�������ַ�����</strong></td>" & vbCrLf
    Response.Write "      <td><input name='UserNameMax' type='text' id='UserNameMax' value='" & rsConfig("UserNameMax") & "' size='6' maxlength='5'> ���ַ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��ֹע����û�����</strong><br>���ұ�ָ�����û���������ֹע�ᣬÿ���û������á�|�����ŷָ�</td>" & vbCrLf
    Response.Write "      <td><input type='text' name='UserName_RegDisabled' value='" & rsConfig("UserName_RegDisabled") & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Աע��ʱ�ı�����Ŀ��</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <table width='100%'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='UserName' checked disabled>�û���</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Password' checked disabled>����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Question' checked disabled>��������</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Answer' checked disabled>�����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Email' checked disabled>Email</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Homepage'" & IsMustFill(RegFields_MustFill, "Homepage") & ">��ҳ</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='QQ'" & IsMustFill(RegFields_MustFill, "QQ") & ">QQ����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='ICQ'" & IsMustFill(RegFields_MustFill, "ICQ") & ">ICQ����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='MSN'" & IsMustFill(RegFields_MustFill, "MSN") & ">MSN�ʺ�</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='UC'" & IsMustFill(RegFields_MustFill, "UC") & ">UC����</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='OfficePhone'" & IsMustFill(RegFields_MustFill, "OfficePhone") & ">�칫�绰</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='HomePhone'" & IsMustFill(RegFields_MustFill, "HomePhone") & ">��ͥ�绰</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Mobile'" & IsMustFill(RegFields_MustFill, "Mobile") & ">�ֻ�����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Fax'" & IsMustFill(RegFields_MustFill, "Fax") & ">�������</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='PHS'" & IsMustFill(RegFields_MustFill, "PHS") & ">С��ͨ</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td colspan='2'><input name='RegFields_MustFill' type='checkbox' value='Region'" & IsMustFill(RegFields_MustFill, "Region") & ">����/������ʡ��/�ݿ�������</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Address'" & IsMustFill(RegFields_MustFill, "Address") & ">��ϵ��ַ</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='ZipCode'" & IsMustFill(RegFields_MustFill, "ZipCode") & ">��������</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Yahoo'" & IsMustFill(RegFields_MustFill, "Yahoo") & ">�Ż�ͨ�ʺ�</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='TrueName'" & IsMustFill(RegFields_MustFill, "TrueName") & ">��ʵ����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Birthday'" & IsMustFill(RegFields_MustFill, "Birthday") & ">��������</td>" & vbCrLf
    'Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Vocation'" & IsMustFill(RegFields_MustFill, "Vocation") & ">ְҵ</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='IDCard'" & IsMustFill(RegFields_MustFill, "IDCard") & ">���֤����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Aim'" & IsMustFill(RegFields_MustFill, "Aim") & ">Aim�ʺ�</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Company'" & IsMustFill(RegFields_MustFill, "Company") & ">��˾/��λ</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Department'" & IsMustFill(RegFields_MustFill, "Department") & ">����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='PosTitle'" & IsMustFill(RegFields_MustFill, "PosTitle") & ">ְ��</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Marriage'" & IsMustFill(RegFields_MustFill, "Marriage") & ">����״��</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Income'" & IsMustFill(RegFields_MustFill, "Income") & ">�������</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='UserFace'" & IsMustFill(RegFields_MustFill, "UserFace") & ">�û�ͷ��</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='FaceWidth'" & IsMustFill(RegFields_MustFill, "FaceWidth") & ">ͷ����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='FaceHeight'" & IsMustFill(RegFields_MustFill, "FaceHeight") & ">ͷ��߶�</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Sign'" & IsMustFill(RegFields_MustFill, "Sign") & ">ǩ����</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Privacy'" & IsMustFill(RegFields_MustFill, "Privacy") & ">��˽�趨</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע��ʱ���͵Ļ��֣�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentExp' type='text' id='PresentExp' value='" & rsConfig("PresentExp") & "' size='6' maxlength='5'> �ֻ���</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע��ʱ���͵Ľ�Ǯ��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentMoney' type='text' id='PresentMoney' value='" & rsConfig("PresentMoney") & "' size='6' maxlength='5'> Ԫ����ң�Ϊ0ʱ�����ͣ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע��ʱ���͵ĵ�����</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentPoint' type='text' id='PresentPoint' value='" & rsConfig("PresentPoint") & "' size='6' maxlength='5'> ���ȯ��Ϊ0ʱ�����ͣ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע��ʱ���͵���Ч�ڣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentValidNum' type='text' id='PresentValidNum' value='" & rsConfig("PresentValidNum") & "' size='6' maxlength='5'>"
    
    Response.Write "      <select name='PresentValidUnit' id='PresentValidUnit'><option value='1' "
    If rsConfig("PresentValidUnit") = 1 Then Response.Write " selected"
    Response.Write ">��</option><option value='2' "
    If rsConfig("PresentValidUnit") = 2 Then Response.Write " selected"
    Response.Write ">��</option><option value='3' "
    If rsConfig("PresentValidUnit") = 3 Then Response.Write " selected"
    Response.Write ">��</option></select>"
    
    Response.Write "��Ϊ0ʱ�����ͣ�Ϊ��1��ʾ�����ڣ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ա��¼ʱ�Ƿ�������֤�빦�ܣ�</strong><br>������֤�빦�ܿ�����һ���̶��Ϸ�ֹ��Ա���뱻�����ƽ�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableCheckCodeOfLogin' value='1' " & IsRadioChecked(rsConfig("EnableCheckCodeOfLogin"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableCheckCodeOfLogin' value='0' " & IsRadioChecked(rsConfig("EnableCheckCodeOfLogin"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Աÿ��¼һ�ν����Ļ��֣�</strong><br>һ��ֻ����һ��</td>" & vbCrLf
    Response.Write "      <td><input name='PresentExpPerLogin' type='text' id='PresentExpPerLogin' value='" & rsConfig("PresentExpPerLogin") & "' size='6' maxlength='5'> �ֻ���</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ա���ʽ����ȯ�Ķһ����ʣ�</strong></td>" & vbCrLf
    Response.Write "      <td>ÿ <input name='MoneyExchangePoint' type='text' id='MoneyExchangePoint' value='" & FormatNumber(rsConfig("MoneyExchangePoint"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> ԪǮ�ɶһ� <strong><font color='#FF0000'>1</font></strong> ���ȯ</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ա���ʽ�����Ч�ڵĶһ����ʣ�</strong></td>" & vbCrLf
    Response.Write "      <td>ÿ <input name='MoneyExchangeValidDay' type='text' id='MoneyExchangeValidDay' value='" & FormatNumber(rsConfig("MoneyExchangeValidDay"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> ԪǮ�ɶһ� <strong><font color='#FF0000'>1</font></strong> ����Ч��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ա�Ļ������ȯ�Ķһ����ʣ�</strong></td>" & vbCrLf
    Response.Write "      <td>ÿ <input name='UserExpExchangePoint' type='text' id='UserExpExchangePoint' value='" & FormatNumber(rsConfig("UserExpExchangePoint"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> �ֻ��ֿɶһ� <strong><font color='#FF0000'>1</font></strong> ���ȯ</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ա�Ļ�������Ч�ڵĶһ����ʣ�</strong></td>" & vbCrLf
    Response.Write "      <td>ÿ <input name='UserExpExchangeValidDay' type='text' id='UserExpExchangeValidDay' value='" & FormatNumber(rsConfig("UserExpExchangeValidDay"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> �ֻ��ֿɶһ� <strong><font color='#FF0000'>1</font></strong> ����Ч��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��ȯ�����ƣ�</strong><br>���磺���ױҡ���ȯ�����</td>" & vbCrLf
    Response.Write "      <td><input name='PointName' type='text' id='PointName' value='" & rsConfig("PointName") & "' size='6' maxlength='5'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��ȯ�ĵ�λ��</strong>���磺�㡢��</td>" & vbCrLf
    Response.Write "      <td><input name='PointUnit' type='text' id='PointUnit' value='" & rsConfig("PointUnit") & "' size='6' maxlength='5'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�»�Աע��ʱ���͵���֤�ʼ����ݣ�</strong><br>�ʼ�����֧��HTML<br><font color='red'>��ǩ˵����</font><br>{$CheckNum}����֤��<br>{$CheckUrl}����Աע����֤��ַ</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfRegCheck' cols='60' rows='5'>" & rsConfig("EmailOfRegCheck") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�ʼ����������</strong><br>" & vbCrLf
    Response.Write "        ��һ��Ҫѡ����������Ѱ�װ�����(��̵�)<br>" & vbCrLf
    Response.Write "        ������ķ�������֧��(�����)�����������ѡ���ޡ�</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <select name='MailObject' id='MailObject'>" & vbCrLf
    Response.Write "          <option value='0'" & IsOptionSelected(rsConfig("MailObject"), 0) & ">��</option>" & vbCrLf
    Response.Write "          <option value='1'" & IsOptionSelected(rsConfig("MailObject"), 1) & ">Jmail " & ShowInstalled("JMail.SMTPMail") & "</option>" & vbCrLf
    Response.Write "          <option value='2'" & IsOptionSelected(rsConfig("MailObject"), 2) & ">CDONTS " & ShowInstalled("CDONTS.NewMail") & "</option>" & vbCrLf
    Response.Write "          <option value='3'" & IsOptionSelected(rsConfig("MailObject"), 3) & ">ASPEMAIL " & ShowInstalled("Persits.MailSender") & "</option>" & vbCrLf
    Response.Write "          <option value='4'" & IsOptionSelected(rsConfig("MailObject"), 4) & ">WebEasyMail " & ShowInstalled("easymail.MailSend") & "</option>" & vbCrLf
    Response.Write "        </select>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP��������ַ��</strong><br>" & vbCrLf
    Response.Write "        ���������ʼ���SMTP������<br>" & vbCrLf
    Response.Write "        ����㲻����˲������壬����ϵ��Ŀռ��� </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailServer' type='text' id='MailServer' value='" & rsConfig("MailServer") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP��¼�û�����</strong><br>" & vbCrLf
    Response.Write "        ����ķ�������ҪSMTP�����֤ʱ�������ô˲���</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailServerUserName' type='text' id='MailServerUserName' value='" & rsConfig("MailServerUserName") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP��¼���룺</strong><br>" & vbCrLf
    Response.Write "        ����ķ�������ҪSMTP�����֤ʱ�������ô˲��� </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailServerPassWord' type='password' id='MailServerPassWord' value='" & rsConfig("MailServerPassWord") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP������</strong><br>" & vbCrLf
    Response.Write "        ����á�name@domain.com���������û�����¼ʱ����ָ��domain.com</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailDomain' type='text' id='MailDomain' value='" & rsConfig("MailDomain") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��������ͼ�����</strong><br>��һ��Ҫѡ����������Ѱ�װ�����</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <select name='PhotoObject' id='PhotoObject'>" & vbCrLf
    Response.Write "          <option value='0'" & IsOptionSelected(rsConfig("PhotoObject"), 0) & ">��</option>" & vbCrLf
    Response.Write "          <option value='1'" & IsOptionSelected(rsConfig("PhotoObject"), 1) & ">AspJpeg��� " & ShowInstalled("Persits.Jpeg") & "</option>" & vbCrLf
    'Response.Write "          <option value='2'" & IsOptionSelected(rsConfig("PhotoObject"), 2) & ">SA-ImgWriter��� " & ShowInstalled("SoftArtisans.ImageGen") & "</option>" & vbCrLf
    'Response.Write "          <option value='3'" & IsOptionSelected(rsConfig("PhotoObject"), 3) & ">SJCatSoft V2.6��� " & ShowInstalled("sjCatSoft.Thumbnail") & "</option>" & vbCrLf
    Response.Write "        </select>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����ͼĬ�Ͽ�ȣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Thumb_DefaultWidth' type='text' value='" & rsConfig("Thumb_DefaultWidth") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> ����&nbsp;&nbsp;&nbsp;&nbsp;��Ϊ0ʱ�����Ը߶�Ϊ׼��������С��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����ͼĬ�ϸ߶ȣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Thumb_DefaultHeight' type='text' value='" & rsConfig("Thumb_DefaultHeight") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> ����&nbsp;&nbsp;&nbsp;&nbsp;��Ϊ0ʱ�����Կ��Ϊ׼��������С��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����ͼ�㷨��</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='Thumb_Arithmetic' value='0' " & IsRadioChecked(rsConfig("Thumb_Arithmetic"), 0) & "> �����㷨����Ⱥ͸߶ȶ�����0ʱ��ֱ����С��ָ����С������һ��Ϊ0ʱ����������С<br>" & vbCrLf
    Response.Write "        <input type='radio' name='Thumb_Arithmetic' value='1' " & IsRadioChecked(rsConfig("Thumb_Arithmetic"), 1) & "> �ü�������Ⱥ͸߶ȶ�����0ʱ���Ȱ���ѱ�����С�ٲü���ָ����С������һ��Ϊ0ʱ����������С��<br>" & vbCrLf
    Response.Write "        <input type='radio' name='Thumb_Arithmetic' value='2' " & IsRadioChecked(rsConfig("Thumb_Arithmetic"), 2) & "> ���䷨����ָ����С�ı���ͼ�ϸ����ϰ���ѱ�����С��ͼƬ��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='ThumbBackgroundColor' " & ISdisplay(rsConfig("Thumb_Arithmetic"), 2) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����ͼ��ɫ��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Thumb_BackgroundColor' type='text' value='" & rsConfig("Thumb_BackgroundColor") & "' size='10' maxlength='10'><img border=0 src='../Editor/images/rect.gif' width=18 style='cursor:hand;backgroundColor:" & rsConfig("Thumb_BackgroundColor") & "' id=s_bordercolor onClick='SelectColor(this,Thumb_BackgroundColor)'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ͼ��������</strong><br>����ͼ����ˮӡ���ͼ������</td>" & vbCrLf
    Response.Write "      <td><input name='PhotoQuality' type='text' value='" & rsConfig("PhotoQuality") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> &nbsp;&nbsp;&nbsp;&nbsp;������50��100������֡�����Խ��ͼ������Խ�á�������Ϊ90��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ˮӡ���ͣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='Watermark_Type' value='0'  " & IsRadioChecked(rsConfig("Watermark_Type"), 0) & " onClick=""PE_Watermark_Text.style.display='';PE_Watermark_Text_FontName.style.display='';PE_Watermark_Text_FontSize.style.display='';PE_Watermark_Text_FontColor.style.display='';PE_Watermark_Text_Bold.style.display='';PE_Watermark_Images_FileName.style.display='none';PE_Watermark_Images_Transparence.style.display='none';PE_Watermark_Images_BackgroundColor.style.display='none'"" > ����ˮӡ&nbsp;&nbsp;"
    Response.Write "          <input type='radio' name='Watermark_Type' value='1'  " & IsRadioChecked(rsConfig("Watermark_Type"), 1) & " onClick=""PE_Watermark_Text.style.display='none';PE_Watermark_Text_FontName.style.display='none';PE_Watermark_Text_FontSize.style.display='none';PE_Watermark_Text_FontColor.style.display='none';PE_Watermark_Text_Bold.style.display='none';PE_Watermark_Images_FileName.style.display='';PE_Watermark_Images_Transparence.style.display='';PE_Watermark_Images_BackgroundColor.style.display=''"" > ͼƬˮӡ</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ˮӡ���֣�</strong><br>ˮӡ�����������˳���15���ַ�����֧���κ�WEB������</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Text' type='text' value='" & rsConfig("Watermark_Text") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_FontName' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�������壺</strong></td>" & vbCrLf
    Response.Write "      <td>"
    Response.Write "        <SELECT name=""Watermark_Text_FontName"" >" & vbCrLf
    Response.Write "            <option value=""����"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "����") & ">����</option>" & vbCrLf
    Response.Write "            <option value=""����_GB2312"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "����_GB2312") & ">����</option>" & vbCrLf
    Response.Write "            <option value=""����_GB2312"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "����_GB2312") & ">������</option>" & vbCrLf
    Response.Write "            <option value=""����"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "����") & ">����</option>" & vbCrLf
    Response.Write "            <option value=""����"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "����") & ">����</option>" & vbCrLf
    Response.Write "            <option value=""��Բ"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "��Բ") & ">��Բ</option>" & vbCrLf
    Response.Write "            <option value=""Andale Mono"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Andale Mono") & ">Andale Mono</OPTION> " & vbCrLf
    Response.Write "            <option value=""Arial""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Arial") & ">Arial</OPTION> " & vbCrLf
    Response.Write "            <option value=""Arial Black""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Arial Black") & ">Arial Black</OPTION> " & vbCrLf
    Response.Write "            <option value=""Book Antiqua""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Book Antiqua") & ">Book Antiqua</OPTION>" & vbCrLf
    Response.Write "            <option value=""Century Gothic""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Century Gothic") & ">Century Gothic</OPTION> " & vbCrLf
    Response.Write "            <option value=""Comic Sans MS""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Comic Sans MS") & ">Comic Sans MS</OPTION>" & vbCrLf
    Response.Write "            <option value=""Courier New""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Courier New") & ">Courier New</OPTION>" & vbCrLf
    Response.Write "            <option value=""Georgia""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Georgia") & ">Georgia</OPTION>" & vbCrLf
    Response.Write "            <option value=""Impact""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Impact") & ">Impact</OPTION>" & vbCrLf
    Response.Write "            <option value=""Tahoma""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Tahoma") & ">Tahoma</OPTION>" & vbCrLf
    Response.Write "            <option value=""Times New Roman""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Times New Roman") & ">Times New Roman</OPTION>" & vbCrLf
    Response.Write "            <option value=""Trebuchet MS""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Trebuchet MS") & ">Trebuchet MS</OPTION>" & vbCrLf
    Response.Write "            <option value=""Script MT Bold""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Script MT Bold") & ">Script MT Bold</OPTION>" & vbCrLf
    Response.Write "            <option value=""Stencil""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Stencil") & ">Stencil</OPTION>" & vbCrLf
    Response.Write "            <option value=""Verdana""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Verdana") & ">Verdana</OPTION>" & vbCrLf
    Response.Write "            <option value=""Lucida Console""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Lucida Console") & ">Lucida Console</OPTION>" & vbCrLf
    Response.Write "        </SELECT>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_FontSize' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>���ִ�С��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Text_FontSize' type='text' value='" & rsConfig("Watermark_Text_FontSize") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> ����</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_FontColor' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>������ɫ��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Text_FontColor' type='text' value='" & rsConfig("Watermark_Text_FontColor") & "' size='10' maxlength='10'><img border=0 src='../Editor/images/rect.gif' width=18 style='cursor:hand;backgroundColor:" & rsConfig("Watermark_Text_FontColor") & "' id=s_bordercolor onClick='SelectColor(this,Watermark_Text_FontColor)'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_Bold' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ���壺</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "          <SELECT name='Watermark_Text_Bold' >" & vbCrLf
    Response.Write "            <OPTION value='0'  " & IsOptionSelected(rsConfig("Watermark_Text_Bold"), False) & ">��</OPTION>" & vbCrLf
    Response.Write "            <OPTION value='1'  " & IsOptionSelected(rsConfig("Watermark_Text_Bold"), True) & ">��</OPTION>" & vbCrLf
    Response.Write "          </SELECT>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Images_FileName' " & ISdisplay(rsConfig("Watermark_Type"), 1) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ˮӡͼƬ�ļ�����</strong><br>��������дͼƬ�ļ������·�����ԡ�\����ͷ</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Images_FileName' type='text' value='" & rsConfig("Watermark_Images_FileName") & "' size='40' maxlength='40'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Images_Transparence' " & ISdisplay(rsConfig("Watermark_Type"), 1) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ͼƬ͸���ȣ�</strong><br> 100% Ϊ��͸��</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Images_Transparence' type='text' value='" & rsConfig("Watermark_Images_Transparence") & "' size='3' maxlength='3'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()"">%</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Images_BackgroundColor' " & ISdisplay(rsConfig("Watermark_Type"), 1) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ͼƬ��ɫ��</strong><br>����ȥ��ˮӡͼƬ�ĵ�ɫ�����ڴ������ɫ��RGBֵ��</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Images_BackgroundColor' type='text' value='" & rsConfig("Watermark_Images_BackgroundColor") & "' size='10' maxlength='10'><img border=0 src='../Editor/images/rect.gif' width=18 style='cursor:hand;backgroundColor:" & rsConfig("Watermark_Images_BackgroundColor") & "' id=s_bordercolor onClick='SelectColor(this,Watermark_Images_BackgroundColor)'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�������λ�ã�</strong></td>" & vbCrLf
    Response.Write "      <td>"
    Response.Write "        <SELECT NAME='Watermark_Position' >" & vbCrLf
    Response.Write "            <option value='0' " & IsOptionSelected(rsConfig("Watermark_Position"), 0) & ">����</option>" & vbCrLf
    Response.Write "            <option value='1' " & IsOptionSelected(rsConfig("Watermark_Position"), 1) & ">����</option>" & vbCrLf
    Response.Write "            <option value='2' " & IsOptionSelected(rsConfig("Watermark_Position"), 2) & ">����</option>" & vbCrLf
    Response.Write "            <option value='3' " & IsOptionSelected(rsConfig("Watermark_Position"), 3) & ">����</option>" & vbCrLf
    Response.Write "            <option value='4' " & IsOptionSelected(rsConfig("Watermark_Position"), 4) & ">����</option>" & vbCrLf
    Response.Write "        </SELECT>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����λ�ã�&nbsp;</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>X��<input name='Watermark_Position_X' type='text' value='" & rsConfig("Watermark_Position_X") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> ����<br>Y��<input name='Watermark_Position_Y' type='text' value='" & rsConfig("Watermark_Position_Y") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> ����</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ÿ������ʱ����</strong>��<br>" & vbCrLf
    Response.Write "        ���ú����ÿ������ʱ���������Ա���������������Ĵ���ϵͳ��Դ</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='SearchInterval' type='text' id='SearchInterval' value='" & rsConfig("SearchInterval") & "' size='10' maxlength='10'> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�����������Ľ����</strong>��<br>" & vbCrLf
    Response.Write "        ���������Ľ���������ĵ���Դ�����ȣ���������ã����鲻Ҫ���ù���</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='SearchResultNum' type='text' id='SearchResultNum' value='" & rsConfig("SearchResultNum") & "' size='10' maxlength='10'> ����¼" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ͨ������ҳ��ÿҳ��Ϣ��</strong>��</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='" & rsConfig("MaxPerPage_SearchResult") & "' size='10' maxlength='10'> ����¼/ҳ" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ�����ȫ������</strong><br>" & vbCrLf
    Response.Write "        ACCESS���ݿⲻ���鿪��<BR>SQL���ݿ�����ȫ���������Կ���" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='SearchContent' value='1' " & IsRadioChecked(rsConfig("SearchContent"), True) & "> ����&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='SearchContent' value='0' " & IsRadioChecked(rsConfig("SearchContent"), False) & "> ����" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ������ο͹�����Ʒ��</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableGuestBuy' value='1' " & IsRadioChecked(rsConfig("EnableGuestBuy"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableGuestBuy' value='0' " & IsRadioChecked(rsConfig("EnableGuestBuy"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ʒ�۸��Ƿ�˰��</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='IncludeTax' value='1' " & IsRadioChecked(rsConfig("IncludeTax"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='IncludeTax' value='0' " & IsRadioChecked(rsConfig("IncludeTax"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>˰�����ã�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='TaxRate' type='text' id='TaxRate' value='" & rsConfig("TaxRate") & "'  size='6' maxlength='6' style='text-align:center'>%</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>����֧��ƽ̨��</strong></td>" & vbCrLf
'    Response.Write "      <td><select name='PayOnlineProvider' id='PayOnlineProvider'>" & vbCrLf
'    Response.Write "          <option value='0'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 0) & ">��</option>" & vbCrLf
'    Response.Write "          <option value='1'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 1) & ">��������1.1��</option>" & vbCrLf
'    Response.Write "          <option value='2'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 2) & ">�й�����֧����</option>" & vbCrLf
'    Response.Write "          <option value='3'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 3) & ">�Ϻ���ѸIPS</option>" & vbCrLf
'    Response.Write "          <option value='4'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 4) & ">�㶫����</option>" & vbCrLf
'    Response.Write "          <option value='5'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 5) & ">����֧��</option>" & vbCrLf
'    Response.Write "          <option value='6'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 6) & ">�׸�ͨ</option>" & vbCrLf
'    Response.Write "          <option value='7'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 7) & ">����֧��</option>" & vbCrLf
'    Response.Write "          <option value='8'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 8) & ">֧����֧��</option>" & vbCrLf
'    Response.Write "          <option value='9'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 9) & ">��Ǯ֧��</option>" & vbCrLf
'    Response.Write "          <option value='10'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 10) & ">��������2.0��</option>" & vbCrLf
'    Response.Write "          <option value='11'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 11) & ">��Ǯ������</option>" & vbCrLf
'    Response.Write "          <option value='13'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 13) & ">�Ƹ�ͨ</option>" & vbCrLf
'    Response.Write "        </select>" & vbCrLf
'    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='http://www.powereasy.net/payreg.html' target='_blank'>���ע�����������̻�</a>"
'    Response.Write "      </td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>�̻���ţ�</strong><br>������������������֧��ƽ̨������̻����</td>" & vbCrLf
'    Response.Write "      <td><input name='PayOnlineShopID' type='text' id='PayOnlineShopID' value='" & rsConfig("PayOnlineShopID") & "' size='30' maxlength='50'></td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>MD5˽Կ��</strong><br>������������������֧��ƽ̨�����õ�MD5˽Կ<br>��������֧��ƽ̨����Ҫ����</td>" & vbCrLf
'    Response.Write "      <td><input name='PayOnlineKey' type='password' id='PayOnlineKey' value='" & rsConfig("PayOnlineKey") & "' size='30' maxlength='255'></td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>�������ʣ�</strong></td>" & vbCrLf
'    Response.Write "      <td>" & vbCrLf
'    Response.Write "        <input name='PayOnlineRate' type='text' id='PayOnlineRate' value='" & rsConfig("PayOnlineRate") & "' size='6' maxlength='6' style='text-align:center'>%<br>" & vbCrLf
'    Response.Write "        <input name='PayOnlinePlusPoundage' type='checkbox' value='1' " & IsRadioChecked(rsConfig("PayOnlinePlusPoundage"), True) & "> �������ɸ����˶���֧��" & vbCrLf
'    Response.Write "      </td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�������ǰ׺��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Prefix_OrderFormNum' type='text' id='Prefix_OrderFormNum' value='" & rsConfig("Prefix_OrderFormNum") & "' size='6' maxlength='4'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����֧�������ǰ׺��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Prefix_PaymentNum' type='text' id='Prefix_PaymentNum' value='" & rsConfig("Prefix_PaymentNum") & "' size='6' maxlength='4'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�����ڵĹ��ң�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Country' type='text' id='Country' value='" & rsConfig("Country") & "' size='15' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�����ڵ�ʡ�ݣ�</strong></td>" & vbCrLf
    Response.Write "      <td><select name='Province'>" & GetProvince(rsConfig("Province")) & "</select></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�����ڵĳ��л������</strong></td>" & vbCrLf
    Response.Write "      <td><input name='City' type='text' id='City' value='" & rsConfig("City") & "' size='15' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�����ڵ������������룺</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PostCode' type='text' id='PostCode' value='" & rsConfig("PostCode") & "' size='10' maxlength='10'> <font color='red'>�����Զ����㶩�����˷�</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ȷ�϶���ʱվ�ڶ���/Email֪ͨ���ݣ�</strong><br>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfOrderConfirm' cols='60' rows='4'>" & rsConfig("EmailOfOrderConfirm") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�յ����л���վ�ڶ���/Email֪ͨ���ݣ�</strong><br>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfReceiptMoney' cols='60' rows='4'>" & rsConfig("EmailOfReceiptMoney") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�˿��վ�ڶ���/Email֪ͨ���ݣ�</strong><br>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfRefund' cols='60' rows='4'>" & rsConfig("EmailOfRefund") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ʊ��վ�ڶ���/Email֪ͨ���ݣ�</strong><br>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfInvoice' cols='60' rows='4'>" & rsConfig("EmailOfInvoice") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>���������վ�ڶ���/Email֪ͨ���ݣ�</strong><br>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfDeliver' cols='60' rows='4'>" & rsConfig("EmailOfDeliver") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>���Ϳ��ź�վ�ڶ���/Email֪ͨ���ݣ�</strong><br>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��<br>�ر��ǩ��<br>{$CardInfo}������Ŀ��ż�������Ϣ</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfSendCard' cols='60' rows='4'>" & rsConfig("EmailOfSendCard") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>֪ͨ�����еĿ��ñ�ǩ�����壺</strong></td>" & vbCrLf
    Response.Write "      <td><textarea name='Labels' cols='60' rows='4' ReadOnly>"
    Response.Write "{$OrderFormID}������ID" & vbCrLf
    Response.Write "{$OrderFormNum}���������" & vbCrLf
    Response.Write "{$ContacterName}���ջ�������" & vbCrLf
    Response.Write "{$OrderInfo}��������Ϣ" & vbCrLf
    Response.Write "{$MoneyTotal}�������ܽ��" & vbCrLf
    Response.Write "{$MoneyReceipt}���������տ�" & vbCrLf
    Response.Write "{$MoneyNeedPay}����Ҫ֧�����" & vbCrLf
    Response.Write "{$InputTime}������ʱ��" & vbCrLf
    Response.Write "{$UserName}����Ա�û���" & vbCrLf
    Response.Write "{$Address}���ջ��˵�ַ" & vbCrLf
    Response.Write "{$ZipCode}���ջ����ʱ�" & vbCrLf
    Response.Write "{$Mobile}���ջ����ֻ�" & vbCrLf
    Response.Write "{$Phone}���ջ��˵绰" & vbCrLf
    Response.Write "{$Email}���ջ���Email" & vbCrLf
    Response.Write "{$PaymentType}�����ʽ" & vbCrLf
    Response.Write "{$DeliverType}���ͻ���ʽ" & vbCrLf
    Response.Write "{$OrderStatus}������״̬" & vbCrLf
    Response.Write "{$PayStatus}���������" & vbCrLf
    Response.Write "{$DeliverStatus}������״̬" & vbCrLf
    Response.Write "{$Charge_Deliver}���˷�" & vbCrLf
    Response.Write "{$PresentMoney}�������ֽ�ȯ" & vbCrLf
    Response.Write "{$PresentExp}�����ͻ���" & vbCrLf
    Response.Write "{$Charge_Deliver}���˷�" & vbCrLf
    Response.Write "{$ExpressCompany}��������˾����" & vbCrLf
    Response.Write "{$ExpressNumber}����ݵ���" & vbCrLf
    Response.Write "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ������ο����ԣ�</strong><br>��ѡ������οͻ�δ��¼�û�����ǩд����</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_EnableVisitor' value='1' " & IsRadioChecked(rsConfig("GuestBook_EnableVisitor"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_EnableVisitor' value='0' " & IsRadioChecked(rsConfig("GuestBook_EnableVisitor"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ��������԰���֤�빦�ܣ�</strong><br>�û�ǩд����ʱ��Ҫ��дϵͳ������ɵ���֤�룬�˹���������Ԥ�����˶���Ⱥ����������</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBookCheck' value='1' " & IsRadioChecked(rsConfig("GuestBookCheck"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBookCheck' value='0' " & IsRadioChecked(rsConfig("GuestBookCheck"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ񹫿��ÿ�IP��</strong><br>��ѡ���ǣ�������߿��Կ��������˵������Ϣ</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_ShowIP' value='1' " & IsRadioChecked(rsConfig("GuestBook_ShowIP"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_ShowIP' value='0' " & IsRadioChecked(rsConfig("GuestBook_ShowIP"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>������ָ����������ԣ�</strong><br>��ѡ���ڷ��������˿�������ǩд�����Բ������κ����</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_IsAssignSort' value='1' " & IsRadioChecked(rsConfig("GuestBook_IsAssignSort"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_IsAssignSort' value='0' " & IsRadioChecked(rsConfig("GuestBook_IsAssignSort"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ���������������湦�ܣ�</strong><br>��ѡ���ǣ�����Ҫ���εĹؼ��ֵ����ò���Ч��</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_EnableManageRubbish' value='1' " & IsRadioChecked(rsConfig("GuestBook_EnableManageRubbish"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_EnableManageRubbish' value='0' " & IsRadioChecked(rsConfig("GuestBook_EnableManageRubbish"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "   <tr class='tdbg'>     " & vbCrLf
    Response.Write "     <td width='40%' class='tdbg5'>               <strong>Ҫ���εĹؼ���</strong>��<br>" & vbCrLf
    Response.Write "      (ע����Ӷ�����ƹؼ��֣����ûس��ָ���<br>����û��ύ�����������к���Ҫ���εĹؼ��֣������ʾ��ֹ���ԣ�)<br> </td> " & vbCrLf
    Response.Write "     <td><textarea name='LockRubbish' cols='50' rows='8' id='LockRubbish'>" & vbCrLf
    Dim rsLockRubbish, arrLockRubbish
    rsLockRubbish = Trim(rsConfig("GuestBook_ManageRubbish"))

    If InStr(rsLockRubbish, "$$$") > 0 Then
        arrLockRubbish = Split(Trim(rsLockRubbish), "$$$")
        For i = 0 To UBound(arrLockRubbish)
            Response.Write arrLockRubbish(i)
            If i < UBound(arrLockRubbish) Then Response.Write Chr(10)
        Next
    Else
        Response.Write rsLockRubbish
    End If

    Response.Write "</textarea>" & vbCrLf
    Response.Write "     </td>    </tr>   " & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Dim GuestBook_MaxPerPage, MaxPerPage
    GuestBook_MaxPerPage = Array(20, 8, 6, 5)
    If Trim(rsConfig("GuestBook_MaxPerPage")) <> "" And Not IsNull(rsConfig("GuestBook_MaxPerPage")) Then
        MaxPerPage = Split(Trim(rsConfig("GuestBook_MaxPerPage")), "|||")
        If UBound(MaxPerPage) = 3 Then GuestBook_MaxPerPage = MaxPerPage
    End If
    Response.Write "      <td width='40%' class='tdbg5'><strong>������������ʽÿҳ��ʾ��������</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_DiscussionMaxPerPage' type='text' id='GuestBook_DiscussionMaxPerPage' value='" & GuestBook_MaxPerPage(0) & "' size='6' maxlength='5'>  ��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�������Ա���ʽÿҳ��ʾ��������</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_GuestBookMaxPerPage' type='text' id='GuestBook_GuestBookMaxPerPage' value='" & GuestBook_MaxPerPage(1) & "' size='6' maxlength='5'>  ��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>���Իظ�ҳÿҳ��ʾ��������</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_ReplyMaxPerPage' type='text' id='GuestBook_ReplyMaxPerPage' value='" & GuestBook_MaxPerPage(2) & "' size='6' maxlength='5'>  ��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����չ����ÿҳ��ʾ��������</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_TreeMaxPerPage' type='text' id='GuestBook_TreeMaxPerPage' value='" & GuestBook_MaxPerPage(3) & "' size='6' maxlength='5'>  ��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ�Ƿ�����RSS���ܣ�</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableRss' value='1' " & IsRadioChecked(rsConfig("EnableRss"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableRss' value='0' " & IsRadioChecked(rsConfig("EnableRss"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='RssSetting' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>Rssʹ�ñ��룺</strong>��<br>" & vbCrLf
    Response.Write "        Rssʹ�õĺ��ֱ���</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='RssCodeType' value='1' " & IsRadioChecked(rsConfig("RssCodeType"), True) & "> GB2312&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='RssCodeType' value='0' " & IsRadioChecked(rsConfig("RssCodeType"), False) & "> UTF-8" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ�Ƿ�����WAP(�ֻ����ʣ����ܣ�</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableWap' value='1' " & IsRadioChecked(rsConfig("EnableWap"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableWap' value='0' " & IsRadioChecked(rsConfig("EnableWap"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>WAP���LOGO</strong>��<br>" & vbCrLf
    Response.Write "        ʹ���ֻ����ʱ��ʾ����վLOGO��<br>����ʹ�ô󲿷���ʽ�ֻ���֧�ֵ�WBMP��ʽͼƬ��<br>ʹ�����ָ�ʽ��ͼƬ��Ҫ�ڷ�����������MIME����<br>wbmp&nbsp;image/vnd.wap.wbmp<br>������Ĭ����ʾ��վ����</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='WapLogo' type='text' id='WapLogo' value='" & rsConfig("WapLogo") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting2' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ��������۹���</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableWapPl' value='1' " & IsRadioChecked(rsConfig("EnableWapPl"), True) & "> ����&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='EnableWapPl' value='0' " & IsRadioChecked(rsConfig("EnableWapPl"), False) & "> ����" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting3' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ����ø�����ʾ</strong><br>" & vbCrLf
    Response.Write "        ĳЩ��ʽ�ֻ�����֧�ֲ�ɫͼƬ��ʾ���翼�Ǽ����ԣ�����رգ�Ŀǰ�������ֻ�������رա�" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapAppendix' value='1' " & IsRadioChecked(rsConfig("ShowWapAppendix"), True) & "> ����&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapAppendix' value='0' " & IsRadioChecked(rsConfig("ShowWapAppendix"), False) & "> ����" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting4' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ������ֻ��̳�</strong><br>" & vbCrLf
    Response.Write "        �����ô����ǿ���û�ע��ʱ������д��ϵ��ַ���������룬��ϵ�绰�����" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapShop' value='1' " & IsRadioChecked(rsConfig("ShowWapShop"), True) & "> ����&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapShop' value='0' " & IsRadioChecked(rsConfig("ShowWapShop"), False) & "> ����" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting5' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�Ƿ�����WAP���к�̨����</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapManage' value='1' " & IsRadioChecked(rsConfig("ShowWapManage"), True) & "> ����&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapManage' value='0' " & IsRadioChecked(rsConfig("ShowWapManage"), False) & "> ����" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>���׶���ͨ���û�����</strong><br>���������� ���׶���ͨƽ̨ ע����û���</td>" & vbCrLf
    Response.Write "      <td><input name='SMSUserName' type='text' id='SMSUserName' value='" & rsConfig("SMSUserName") & "' size='30' maxlength='50'> &nbsp;&nbsp;<a href='http://sms.powereasy.net/Register.aspx' target='_blank'><font color='red'>���ע�����û�</font></a> &nbsp;&nbsp;<a href='http://sms.powereasy.net/' target='_blank'><font color='blue'>ʲô�Ƕ��׶���ͨ��</font></a></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>MD5��Կ��</strong><br>���������� ���׶���ͨƽ̨ �����õ�MD5��Կ</td>" & vbCrLf
    Response.Write "      <td><input name='SMSKey' type='password' id='SMSKey' value='" & rsConfig("SMSKey") & "' size='30' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�ͻ��ύ����ʱ��ϵͳ�Ƿ��Զ������ֻ�����֪ͨ����Ա��</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='SendMessageToAdminWhenOrder' value='1' " & IsRadioChecked(rsConfig("SendMessageToAdminWhenOrder"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='SendMessageToAdminWhenOrder' value='0' " & IsRadioChecked(rsConfig("SendMessageToAdminWhenOrder"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա��С��ͨ���ֻ����룺</strong><br>ÿ������һ�����롣<br>�������������룬ϵͳ��ͬʱ���͵����������</td>" & vbCrLf
    Response.Write "      <td><textarea name='Mobiles' cols='60' rows='4'>" & rsConfig("Mobiles") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�ͻ��¶���ʱϵͳ������Ա���Ͷ��ŵ����ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfOrder' cols='60' rows='4'>" & rsConfig("MessageOfOrder") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�ͻ�����֧���ɹ����Ƿ���ͻ������ֻ����ţ���֪�俨�ź����룺</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='SendMessageToMemberWhenPaySuccess' value='1' " & IsRadioChecked(rsConfig("SendMessageToMemberWhenPaySuccess"), True) & "> �� &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='SendMessageToMemberWhenPaySuccess' value='0' " & IsRadioChecked(rsConfig("SendMessageToMemberWhenPaySuccess"), False) & "> ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>ȷ�϶���ʱ�ֻ�����֪ͨ���ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfOrderConfirm' cols='60' rows='4'>" & rsConfig("MessageOfOrderConfirm") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�յ����л����ֻ�����֪ͨ���ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfReceiptMoney' cols='60' rows='4'>" & rsConfig("MessageOfReceiptMoney") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>�˿���ֻ�����֪ͨ���ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfRefund' cols='60' rows='4'>" & rsConfig("MessageOfRefund") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ʊ���ֻ�����֪ͨ���ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfInvoice' cols='60' rows='4'>" & rsConfig("MessageOfInvoice") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����������ֻ�����֪ͨ���ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfDeliver' cols='60' rows='4'>" & rsConfig("MessageOfDeliver") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>���Ϳ��ź��ֻ�����֪ͨ���ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��<br>�ر��ǩ��<br>{$CardInfo}������Ŀ��ż�������Ϣ</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfSendCard' cols='60' rows='4'>" & rsConfig("MessageOfSendCard") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>֪ͨ�����еĿ��ñ�ǩ�����壺</strong></td>" & vbCrLf
    Response.Write "      <td><textarea name='Labels' cols='60' rows='4' ReadOnly>"
    Response.Write "{$OrderFormID}������ID" & vbCrLf
    Response.Write "{$OrderFormNum}���������" & vbCrLf
    Response.Write "{$ContacterName}���ջ�������" & vbCrLf
    Response.Write "{$OrderInfo}��������Ϣ" & vbCrLf
    Response.Write "{$MoneyTotal}�������ܽ��" & vbCrLf
    Response.Write "{$MoneyReceipt}���������տ�" & vbCrLf
    Response.Write "{$MoneyNeedPay}����Ҫ֧�����" & vbCrLf
    Response.Write "{$InputTime}������ʱ��" & vbCrLf
    Response.Write "{$UserName}����Ա�û���" & vbCrLf
    Response.Write "{$Address}���ջ��˵�ַ" & vbCrLf
    Response.Write "{$ZipCode}���ջ����ʱ�" & vbCrLf
    Response.Write "{$Mobile}���ջ����ֻ�" & vbCrLf
    Response.Write "{$Phone}���ջ��˵绰" & vbCrLf
    Response.Write "{$Email}���ջ���Email" & vbCrLf
    Response.Write "{$PaymentType}�����ʽ" & vbCrLf
    Response.Write "{$DeliverType}���ͻ���ʽ" & vbCrLf
    Response.Write "{$OrderStatus}������״̬" & vbCrLf
    Response.Write "{$PayStatus}���������" & vbCrLf
    Response.Write "{$DeliverStatus}������״̬" & vbCrLf
    Response.Write "{$Charge_Deliver}���˷�" & vbCrLf
    Response.Write "{$PresentMoney}�������ֽ�ȯ" & vbCrLf
    Response.Write "{$PresentExp}�����ͻ���" & vbCrLf
    Response.Write "{$Charge_Deliver}���˷�" & vbCrLf
    Response.Write "{$ExpressCompany}��������˾����" & vbCrLf
    Response.Write "{$ExpressNumber}����ݵ���" & vbCrLf
    Response.Write "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'>&nbsp;</td>" & vbCrLf
    Response.Write "      <td> </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա������л���¼ʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$Balance}���ʽ����<br>{$ReceiptDate}����������<br>{$Money}�������<br>{$BankName}����������</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddRemit' cols='60' rows='4'>" & rsConfig("MessageOfAddRemit") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա������������¼ʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$Balance}���ʽ����<br>{$Money}��������<br>{$Reason}��ԭ��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddIncome' cols='60' rows='4'>" & rsConfig("MessageOfAddIncome") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա���֧����¼ʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$Balance}���ʽ����<br>{$Money}��֧�����<br>{$Reason}��ԭ��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddPayment' cols='60' rows='4'>" & rsConfig("MessageOfAddPayment") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա�һ���ȯʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$Balance}���ʽ����<br>{$UserPoint}�����õ�ȯ<br>{$Money}��֧�����<br>{$Point}���õ��ĵ�ȯ��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfExchangePoint' cols='60' rows='4'>" & rsConfig("MessageOfExchangePoint") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա������ȯʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$UserPoint}�����õ�ȯ<br>{$Point}�����ӵĵ�ȯ��<br>{$Reason}������ԭ��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddPoint' cols='60' rows='4'>" & rsConfig("MessageOfAddPoint") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա�۳���ȯʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$UserPoint}�����õ�ȯ<br>{$Point}���۳��ĵ�ȯ��<br>{$Reason}���۳�ԭ��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfMinusPoint' cols='60' rows='4'>" & rsConfig("MessageOfMinusPoint") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա�һ���Ч��ʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$Balance}���ʽ����<br>{$ValidDays}��ʣ������<br>{$Money}��֧�����<br>{$Valid}���õ�����Ч��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfExchangeValid' cols='60' rows='4'>" & rsConfig("MessageOfExchangeValid") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա������Ч��ʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$ValidDays}��ʣ������<br>{$Valid}���õ�����Ч��<br>{$Reason}������ԭ��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddValid' cols='60' rows='4'>" & rsConfig("MessageOfAddValid") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>����Ա�۳���Ч��ʱ���͵��ֻ��������ݣ�</strong><br>��֧��HTML���룬���ñ�ǩ��<br>{$UserName}����Ա�û���<br>{$ValidDays}��ʣ������<br>{$Valid}���۳�����Ч��<br>{$Reason}���۳�ԭ��</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfMinusValid' cols='60' rows='4'>" & rsConfig("MessageOfMinusValid") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    
    Response.Write "<table width='100%' border='0'>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' align='center'>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex_Old' type='hidden' id='FileExt_SiteIndex_Old' value='" & rsConfig("FileExt_SiteIndex") & "'>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial_Old' type='hidden' id='FileExt_SiteSpecial_Old' value='" & rsConfig("FileExt_SiteSpecial") & "'>" & vbCrLf
    Response.Write "        <input name='Modules_Old' type='hidden' id='Modules_Old' value='" & rsConfig("Modules") & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveConfig'>" & vbCrLf
    Response.Write "        <input name='cmdSave' type='submit' id='cmdSave' value=' �������� '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "    <form name='mysitekeyform' id='mysitekeyform' method='POST' action='http://www.powereasy.net/genuine/CheckSite.asp?CheckType=SiteKey' target='_blank'>" & vbCrLf
    Response.Write "  <input type='hidden' id='SiteKey' name='SiteKey' value=''>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    rsConfig.Close
    Set rsConfig = Nothing
End Sub

Sub SaveConfig()
    Dim sqlConfig, rsConfig, iSiteKey, FoundErr
    FoundErr = False

    If Trim(Request("AdminDir")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��̨����Ŀ¼����Ϊ��</li>"
    End If

    If Trim(Request("ADDir")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��վ���Ŀ¼����Ϊ��</li>"
    End If

    If PE_CLng(Trim(Request("PhotoObject"))) > 0 Then
        If PE_CLng(Trim(Request("Watermark_Type"))) = 0 Then
            If Trim(Request("Watermark_Text")) <> "" Then
                If Trim(Request("Watermark_Text_FontColor")) = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��������ˮӡ����ɫ����Ϊ��</li>"
                End If
            End If
        Else
            If Trim(Request("Watermark_Images_FileName")) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>����ˮӡͼƬ·������Ϊ��</li>"
            Else
                If Trim(Request("Watermark_Images_BackgroundColor")) = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����ȥ��ˮӡͼƬ����ɫ����Ϊ��</li>"
                End If
                If Not fso.FileExists(Server.MapPath(Trim(Request("Watermark_Images_FileName")))) Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����ˮӡͼƬ��ͼƬ·������,��ָ��·���е�ͼƬ�����ڡ�</li>"
                End If
            End If
        End If
    End If
    Dim arrLockIP, arrIpW, arrIpB, i, arrLockIPCut
    arrLockIP = Split(Trim(Request("LockIPWhite")), vbCrLf)
    For i = 0 To UBound(arrLockIP)
        If Not (arrLockIP(i) = "" Or IsNull(arrLockIP(i))) And InStr(Trim(arrLockIP(i)), "----") > 0 Then
                arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
                If Not isIP(Trim(arrLockIPCut(0))) Or Not isIP(Trim(arrLockIPCut(1))) Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "����д��ȷ��վ�ڰ������е�IP��ַ��"
                    Exit For
                End If
                If i = 0 Then
                    arrIpW = EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
                Else
                    arrIpW = arrIpW & "$$$" & EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
                End If
        End If
    Next
    arrLockIP = Split(Trim(Request("LockIPBlack")), vbCrLf)
    For i = 0 To UBound(arrLockIP)
        If Not (arrLockIP(i) = "" Or IsNull(arrLockIP(i))) And InStr(Trim(arrLockIP(i)), "----") > 0 Then
            arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
            If Not isIP(Trim(arrLockIPCut(0))) Or Not isIP(Trim(arrLockIPCut(1))) Then
                FoundErr = True
                ErrMsg = ErrMsg & "����д��ȷ��վ�ڰ������е�IP��ַ��"
                Exit For
            End If
            If i = 0 Then
                arrIpB = EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
            Else
                arrIpB = arrIpB & "$$$" & EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
            End If
        End If
    Next

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    sqlConfig = "select * from PE_Config"
    Set rsConfig = Server.CreateObject("ADODB.Recordset")
    rsConfig.Open sqlConfig, Conn, 1, 3

    If rsConfig.BOF And rsConfig.EOF Then
        rsConfig.addnew
    End If

    rsConfig("SiteName") = Trim(Request("SiteName"))
    rsConfig("SiteTitle") = Trim(Request("SiteTitle"))
    rsConfig("SiteUrl") = Trim(Request("SiteUrl"))
    rsConfig("InstallDir") = InstallDir
    rsConfig("LogoUrl") = Trim(Request("LogoUrl"))
    rsConfig("BannerUrl") = Trim(Request("BannerUrl"))
    rsConfig("WebmasterName") = Trim(Request("WebmasterName"))
    rsConfig("WebmasterEmail") = Trim(Request("WebmasterEmail"))
    rsConfig("Copyright") = Trim(Request("Copyright"))
    rsConfig("Meta_Keywords") = Trim(Request("Meta_Keywords"))
    rsConfig("Meta_Description") = Trim(Request("Meta_Description"))
    rsConfig("SiteKey") = Trim(Request("SaveSiteKey"))

    rsConfig("ShowSiteChannel") = PE_CBool(Trim(Request("ShowSiteChannel")))
    rsConfig("ShowAdminLogin") = PE_CBool(Trim(Request("ShowAdminLogin")))
    rsConfig("EnableSaveRemote") = PE_CBool(Trim(Request("EnableSaveRemote")))
    rsConfig("EnableLinkReg") = PE_CBool(Trim(Request("EnableLinkReg")))
    rsConfig("EnableCountFriendSiteHits") = PE_CBool(Trim(Request("EnableCountFriendSiteHits")))
    rsConfig("EnableSoftKey") = PE_CBool(Trim(Request("EnableSoftKey")))
    rsConfig("IsCustom_Content") = PE_CBool(Trim(Request("IsCustom_Content")))
    rsConfig("objName_FSO") = Trim(Request("objName_FSO"))
    rsConfig("AdminDir") = Trim(Request("AdminDir"))
    rsConfig("ADDir") = Trim(Request("ADDir"))
    rsConfig("AnnounceCookieTime") = PE_CLng(Trim(Request("AnnounceCookieTime")))
    rsConfig("HitsOfHot") = PE_CLng(Trim(Request("HitsOfHot")))
    rsConfig("Modules") = ReplaceBadChar(Trim(Request("Modules")))
    rsConfig("FileExt_SiteIndex") = PE_CLng(Trim(Request("FileExt_SiteIndex")))
    rsConfig("FileExt_SiteSpecial") = PE_CLng(Trim(Request("FileExt_SiteSpecial")))
    rsConfig("SiteUrlType") = PE_CLng(Trim(Request("SiteUrlType")))
    'rsConfig("LockIPType") = PE_CLng(Trim(Request("LockIPW"))) + PE_CLng(Trim(Request("LockIPB")))
    rsConfig("LockIPType") = PE_CLng(Trim(Request("LockIPType")))
    rsConfig("LockIP") = arrIpW & "|||" & arrIpB

    rsConfig("EnableUserReg") = PE_CBool(Trim(Request("EnableUserReg")))
    rsConfig("EmailCheckReg") = PE_CBool(Trim(Request("EmailCheckReg")))
    rsConfig("AdminCheckReg") = PE_CBool(Trim(Request("AdminCheckReg")))
    rsConfig("EnableMultiRegPerEmail") = PE_CBool(Trim(Request("EnableMultiRegPerEmail")))
    rsConfig("EnableCheckCodeOfLogin") = PE_CBool(Trim(Request("EnableCheckCodeOfLogin")))
    rsConfig("EnableCheckCodeOfReg") = PE_CBool(Trim(Request("EnableCheckCodeOfReg")))
    rsConfig("EnableQAofReg") = PE_CBool(Trim(Request("EnableQAofReg")))
    rsConfig("QAofReg") = Trim(Request("RegQuestion1")) & " $$$" & Trim(Request("RegAnswer1")) & " $$$" & Trim(Request("RegQuestion2")) & " $$$" & Trim(Request("RegAnswer2")) & " $$$" & Trim(Request("RegQuestion3")) & " $$$" & Trim(Request("RegAnswer3"))

    rsConfig("UserNameLimit") = PE_CLng(Trim(Request("UserNameLimit")))
    rsConfig("UserNameMax") = PE_CLng(Trim(Request("UserNameMax")))
    rsConfig("UserName_RegDisabled") = Trim(Request("UserName_RegDisabled"))
    rsConfig("RegFields_MustFill") = ReplaceBadChar(Trim(Request("RegFields_MustFill")))

    rsConfig("PresentExp") = PE_CLng(Trim(Request("PresentExp")))
    rsConfig("PresentMoney") = PE_CDbl(Trim(Request("PresentMoney")))
    rsConfig("PresentPoint") = PE_CLng(Trim(Request("PresentPoint")))
    rsConfig("PresentValidNum") = PE_CLng(Trim(Request("PresentValidNum")))
    rsConfig("PresentValidUnit") = PE_CLng(Trim(Request("PresentValidUnit")))
    rsConfig("PresentExpPerLogin") = PE_CLng(Trim(Request("PresentExpPerLogin")))
    rsConfig("MoneyExchangePoint") = PE_CDbl(Trim(Request("MoneyExchangePoint")))
    rsConfig("MoneyExchangeValidDay") = PE_CDbl(Trim(Request("MoneyExchangeValidDay")))
    rsConfig("UserExpExchangePoint") = PE_CDbl(Trim(Request("UserExpExchangePoint")))
    rsConfig("UserExpExchangeValidDay") = PE_CDbl(Trim(Request("UserExpExchangeValidDay")))
    rsConfig("PointName") = Trim(Request("PointName"))
    rsConfig("PointUnit") = Trim(Request("PointUnit"))
    rsConfig("EmailOfRegCheck") = Trim(Request("EmailOfRegCheck"))
    rsConfig("ShowAnonymous") = PE_CBool(Trim(Request("ShowAnonymous")))
	
    rsConfig("MailObject") = Trim(Request("MailObject"))
    rsConfig("MailServer") = Trim(Request("MailServer"))
    rsConfig("MailServerUserName") = Trim(Request("MailServerUserName"))
    rsConfig("MailServerPassWord") = Trim(Request("MailServerPassWord"))
    rsConfig("MailDomain") = Trim(Request("MailDomain"))
    
    rsConfig("PhotoObject") = PE_CLng(Trim(Request("PhotoObject")))
    rsConfig("Thumb_DefaultWidth") = PE_CLng(Trim(Request("Thumb_DefaultWidth")))
    rsConfig("Thumb_DefaultHeight") = PE_CLng(Trim(Request("Thumb_DefaultHeight")))
    rsConfig("Thumb_Arithmetic") = PE_CLng(Trim(Request("Thumb_Arithmetic")))
    rsConfig("Thumb_BackgroundColor") = Trim(Request("Thumb_BackgroundColor"))
    rsConfig("PhotoQuality") = PE_CLng(Trim(Request("PhotoQuality")))

    rsConfig("Watermark_Type") = PE_CLng(Trim(Request("Watermark_Type")))
    rsConfig("Watermark_Text") = Trim(Request("Watermark_Text"))
    rsConfig("Watermark_Text_FontName") = Trim(Request("Watermark_Text_FontName"))
    rsConfig("Watermark_Text_FontSize") = PE_CLng(Trim(Request("Watermark_Text_FontSize")))
    rsConfig("Watermark_Text_FontColor") = Trim(Request("Watermark_Text_FontColor"))
    rsConfig("Watermark_Text_Bold") = PE_CBool(Trim(Request("Watermark_Text_Bold")))
    rsConfig("Watermark_Images_FileName") = Trim(Request("Watermark_Images_FileName"))
    rsConfig("Watermark_Images_Transparence") = PE_CLng(Trim(Request("Watermark_Images_Transparence")))
    rsConfig("Watermark_Images_BackgroundColor") = Trim(Request("Watermark_Images_BackgroundColor"))
    rsConfig("Watermark_Position_X") = PE_CLng(Trim(Request("Watermark_Position_X")))
    rsConfig("Watermark_Position_Y") = PE_CLng(Trim(Request("Watermark_Position_Y")))
    rsConfig("Watermark_Position") = PE_CLng(Trim(Request("Watermark_Position")))
    
    rsConfig("SearchInterval") = PE_CLng(Trim(Request("SearchInterval")))
    rsConfig("SearchResultNum") = PE_CLng(Trim(Request("SearchResultNum")))
    rsConfig("MaxPerPage_SearchResult") = PE_CLng(Trim(Request("MaxPerPage_SearchResult")))
    rsConfig("SearchContent") = PE_CBool(Trim(Request("SearchContent")))
    
    rsConfig("EnableGuestBuy") = PE_CBool(Trim(Request("EnableGuestBuy")))
    rsConfig("IncludeTax") = PE_CBool(Trim(Request("IncludeTax")))
    rsConfig("TaxRate") = PE_CLng(Trim(Request("TaxRate")))

'    rsConfig("PayOnlineProvider") = Trim(Request("PayOnlineProvider"))
'    rsConfig("PayOnlineShopID") = Trim(Request("PayOnlineShopID"))
'    rsConfig("PayOnlineKey") = Trim(Request("PayOnlineKey"))
'    rsConfig("PayOnlineRate") = CDbl(Trim(Request("PayOnlineRate")))
'
'    If Trim(Request("PayOnlinePlusPoundage")) = "1" Then
'        rsConfig("PayOnlinePlusPoundage") = True
'    Else
'        rsConfig("PayOnlinePlusPoundage") = False
'    End If

    rsConfig("Prefix_OrderFormNum") = Trim(Request("Prefix_OrderFormNum"))
    rsConfig("Prefix_PaymentNum") = Trim(Request("Prefix_PaymentNum"))

    rsConfig("Country") = Trim(Request("Country"))
    rsConfig("Province") = Trim(Request("Province"))
    rsConfig("City") = Trim(Request("City"))
    rsConfig("PostCode") = Trim(Request("PostCode"))
    rsConfig("EmailOfOrderConfirm") = Trim(Request("EmailOfOrderConfirm"))
    rsConfig("EmailOfSendCard") = Trim(Request("EmailOfSendCard"))
    rsConfig("EmailOfReceiptMoney") = Trim(Request("EmailOfReceiptMoney"))
    rsConfig("EmailOfRefund") = Trim(Request("EmailOfRefund"))
    rsConfig("EmailOfInvoice") = Trim(Request("EmailOfInvoice"))
    rsConfig("EmailOfDeliver") = Trim(Request("EmailOfDeliver"))
    rsConfig("ShowUserModel") = PE_CBool(Trim(Request("ShowUserModel")))	
    
    rsConfig("GuestBook_EnableVisitor") = PE_CBool(Trim(Request("GuestBook_EnableVisitor")))
    rsConfig("GuestBookCheck") = PE_CBool(Trim(Request("GuestBookCheck")))
    rsConfig("GuestBook_EnableManageRubbish") = PE_CBool(Trim(Request("GuestBook_EnableManageRubbish")))
    Dim arrLockRubbish, arrRubbish
    arrLockRubbish = Split(Trim(Request("LockRubbish")), vbCrLf)
    For i = 0 To UBound(arrLockRubbish)
        If Not (arrLockRubbish(i) = "" Or IsNull(arrLockRubbish(i))) Then
            If i = 0 Then
                arrRubbish = Trim(arrLockRubbish(i))
            Else
                arrRubbish = arrRubbish & "$$$" & Trim(arrLockRubbish(i))
            End If
        End If
    Next
    rsConfig("GuestBook_ManageRubbish") = arrRubbish
    rsConfig("GuestBook_ShowIP") = PE_CBool(Trim(Request("GuestBook_ShowIP")))
    rsConfig("GuestBook_IsAssignSort") = PE_CBool(Trim(Request("GuestBook_IsAssignSort")))
    rsConfig("GuestBook_MaxPerPage") = PE_CLng(Trim(Request("GuestBook_DiscussionMaxPerPage"))) & "|||" & PE_CLng(Trim(Request("GuestBook_GuestBookMaxPerPage"))) & "|||" & PE_CLng(Trim(Request("GuestBook_ReplyMaxPerPage"))) & "|||" & PE_CLng(Trim(Request("GuestBook_TreeMaxPerPage")))

    rsConfig("EnableRss") = PE_CBool(Trim(Request("EnableRss")))
    rsConfig("RssCodeType") = PE_CBool(Trim(Request("RssCodeType")))

    rsConfig("EnableWap") = PE_CBool(Trim(Request("EnableWap")))

    If Trim(Request("WapLogo")) = "" Then
        rsConfig("WapLogo") = 0
    Else
        rsConfig("WapLogo") = Trim(Request("WapLogo"))
    End If

    rsConfig("EnableWapPl") = PE_CBool(Trim(Request("EnableWapPl")))
    rsConfig("ShowWapAppendix") = PE_CBool(Trim(Request("ShowWapAppendix")))
    rsConfig("ShowWapShop") = PE_CBool(Trim(Request("ShowWapShop")))
    rsConfig("ShowWapManage") = PE_CBool(Trim(Request("ShowWapManage")))
    
    rsConfig("SMSUserName") = Trim(Request("SMSUserName"))
    rsConfig("SMSKey") = Trim(Request("SMSKey"))
    rsConfig("SendMessageToAdminWhenOrder") = PE_CBool(Trim(Request("SendMessageToAdminWhenOrder")))
    rsConfig("SendMessageToMemberWhenPaySuccess") = PE_CBool(Trim(Request("SendMessageToMemberWhenPaySuccess")))
    rsConfig("Mobiles") = Trim(Request("Mobiles"))
    rsConfig("MessageOfOrder") = Trim(Request("MessageOfOrder"))
    rsConfig("MessageOfOrderConfirm") = Trim(Request("MessageOfOrderConfirm"))
    rsConfig("MessageOfSendCard") = Trim(Request("MessageOfSendCard"))
    rsConfig("MessageOfReceiptMoney") = Trim(Request("MessageOfReceiptMoney"))
    rsConfig("MessageOfRefund") = Trim(Request("MessageOfRefund"))
    rsConfig("MessageOfInvoice") = Trim(Request("MessageOfInvoice"))
    rsConfig("MessageOfDeliver") = Trim(Request("MessageOfDeliver"))

    rsConfig("MessageOfAddRemit") = Trim(Request("MessageOfAddRemit"))
    rsConfig("MessageOfAddIncome") = Trim(Request("MessageOfAddIncome"))
    rsConfig("MessageOfAddPayment") = Trim(Request("MessageOfAddPayment"))
    rsConfig("MessageOfExchangePoint") = Trim(Request("MessageOfExchangePoint"))
    rsConfig("MessageOfAddPoint") = Trim(Request("MessageOfAddPoint"))
    rsConfig("MessageOfMinusPoint") = Trim(Request("MessageOfMinusPoint"))
    rsConfig("MessageOfExchangeValid") = Trim(Request("MessageOfExchangeValid"))
    rsConfig("MessageOfAddValid") = Trim(Request("MessageOfAddValid"))
    rsConfig("MessageOfMinusValid") = Trim(Request("MessageOfMinusValid"))

    rsConfig.Update
    rsConfig.Close
    Set rsConfig = Nothing
    Dim strSql
    If FoundInArr(Request("Modules"), "Supply", ",") Then
        strSql = "Update PE_Channel Set Disabled=" & PE_False & " Where ModuleType=6"
        Conn.Execute (strSql)
    Else
        strSql = "Update PE_Channel Set Disabled=" & PE_True & " Where ModuleType=6"
        Conn.Execute (strSql)
    End If
    If FoundInArr(Request("Modules"), "Job", ",") Then
        strSql = "Update PE_Channel Set Disabled=" & PE_False & " Where ModuleType=8"
        Conn.Execute (strSql)
    Else
        strSql = "Update PE_Channel Set Disabled=" & PE_True & " Where ModuleType=8"
        Conn.Execute (strSql)
    End If
    If FoundInArr(Request("Modules"), "House", ",") Then
        strSql = "Update PE_Channel Set Disabled=" & PE_False & " Where ModuleType=7"
        Conn.Execute (strSql)
    Else
        strSql = "Update PE_Channel Set Disabled=" & PE_True & " Where ModuleType=7"
        Conn.Execute (strSql)
    End If
    
    Call WriteSuccessMsg("��վ���ñ���ɹ���", ComeUrl)

    Dim FileExt_SiteIndex, FileExt_SiteIndex_Old, FileExt_SiteSpecial, FileExt_SiteSpecial_Old
    FileExt_SiteIndex = PE_CLng(Trim(Request("FileExt_SiteIndex")))
    FileExt_SiteIndex_Old = PE_CLng(Trim(Request("FileExt_SiteIndex_Old")))
    FileExt_SiteSpecial = PE_CLng(Trim(Request("FileExt_SiteSpecial")))
    FileExt_SiteSpecial_Old = PE_CLng(Trim(Request("FileExt_SiteSpecial_Old")))
    
    If IsReload(FileExt_SiteIndex, FileExt_SiteIndex_Old) Or IsReload(FileExt_SiteSpecial, FileExt_SiteSpecial_Old) Or Trim(Request("Modules")) <> Trim(Request("Modules_Old")) Then
        Call ReloadLeft
    End If

End Sub

Function IsReload(FileExt, FileExt_Old)
    IsReload = False
    If FileExt <> FileExt_Old Then
        If FileExt = 4 Or FileExt_Old = 4 Then
            IsReload = True
        End If
    End If
End Function

Sub ReloadLeft()
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "  parent.left.location.reload();" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Function ShowInstalled(strObject)
    If Not IsObjInstalled(strObject) Then
        ShowInstalled = "<font color='red'><b>��</b></font>"
    Else
        ShowInstalled = "<b>��</b>"
    End If
End Function

Function IsModulesSelected(Compare1, Compare2)
    If FoundInArr(Compare1, Compare2, ",") = True Then
        IsModulesSelected = " checked"
    Else
        IsModulesSelected = ""
    End If
End Function

Function IsRadioChecked(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If
End Function

Function IsOptionSelected(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If
End Function

Function IsMustFill(Compare1, Compare2)
    If FoundInArr(Compare1, Compare2, ",") = True Then
        IsMustFill = " checked"
    Else
        IsMustFill = ""
    End If
End Function
Function GetProvince(ProvinceName)
    Dim rsProvince, strProvince
    strProvince = "<option value=''>��ѡ��ʡ��</option>"
    Set rsProvince = Conn.Execute("select DISTINCT Province from PE_City")
    Do While Not rsProvince.EOF
        If rsProvince(0) = ProvinceName Then
            strProvince = strProvince & "<option value='" & rsProvince(0) & "' selected>" & rsProvince(0) & "</option>"
        Else
            strProvince = strProvince & "<option value='" & rsProvince(0) & "'>" & rsProvince(0) & "</option>"
        End If
        rsProvince.MoveNext
    Loop
    Set rsProvince = Nothing
    GetProvince = strProvince
End Function

Function ISdisplay(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        ISdisplay = " style='display:'"
    Else
        ISdisplay = " style='display:none'"
    End If
End Function
%>
