<!--#include file="Admin_Common.asp"-->
<!--#include file="RootClass_Menu_Config.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 1   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

If ChannelID = 0 Then
    Response.Write "Ƶ��������"
    Response.End
End If

If AdminPurview > 1 And CheckPurview_Other(AdminPurview_Others, "Menu_" & ChannelDir) = False Then
    Response.Write "��û�д��������Ȩ�ޣ�"
    Response.End
End If

Dim strTopMenu, pNum, pNum2, OpenType_Class, strMenuJS

Response.Write "<html><head><title>������Ŀ�˵�����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Response.Write "  <tr class='topbg'> " & vbCrLf
Response.Write "    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>������Ŀ�˵�����</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=10013' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>"
Response.Write "    <td height='30' colspan='2'>"
Response.Write "<a href='Admin_RootClass_Menu.asp?Action=ShowConfig&ChannelID=" & ChannelID & "' target=main>������Ŀ�˵���������</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_RootClass_Menu.asp?Action=ShowCreate&ChannelID=" & ChannelID & "' target=main>������Ŀ�˵�����</a>"
Response.Write "    </td>"
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>"
Response.Write "    <td width='70' height='30'><strong>�˵���ʾ��</strong></td>"
Response.Write "    <td height='30'>"
Call ShowDemoMenu
Response.Write "    </td>"
Response.Write "    <td width='350'>ע���������ã������������ͣʱЧ��������������Ƴ�ʱЧ����</td>"
Response.Write "  </tr></table>" & vbCrLf

If Action = "ShowConfig" Then
    Call ShowConfig
ElseIf Action = "SaveConfig" Then
    Call SaveConfig
ElseIf Action = "ShowCreate" Then
    Call ShowCreate_RootClass_Menu
ElseIf Action = "Create" Then
    Call Create_RootClass_Menu
Else
    Call ShowConfig
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>" & vbCrLf
Call CloseConn

Sub ShowConfig()
    Response.Write "<form method='POST' action='Admin_RootClass_Menu.asp' id='myform' name='myform'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='6'><strong>������Ŀ�˵���������</strong> ��ע��������Чֻ���ض������������Ч��</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>������ʽ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_1' id='RCM_Menu_1'>"
    Response.Write "        <option value='1' "
    If RCM_Menu_1 = "1" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_1 = "2" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_1 = "3" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_1 = "4" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>����ƫ������</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_2' type='text' id='RCM_Menu_2' value='" & RCM_Menu_2 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>����ƫ������</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_3' type='text' id='RCM_Menu_3' value='" & RCM_Menu_3 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵���߾ࣺ</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_4' type='text' id='RCM_Menu_4' value='" & RCM_Menu_4 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵����ࣺ</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_5' type='text' id='RCM_Menu_5' value='" & RCM_Menu_5 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�����߾ࣺ</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_6' type='text' id='RCM_Menu_6' value='" & RCM_Menu_6 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵����ұ߾ࣺ</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_7' type='text' id='RCM_Menu_7' value='" & RCM_Menu_7 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�͸���ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_8' type='text' id='RCM_Menu_8' value='" & RCM_Menu_8 & "' size='10' maxlength='10' title='0-100 ��ȫ͸��-��ȫ��͸��'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�������Ч��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_9' type='text' id='RCM_Menu_9' value='" & RCM_Menu_9 & "' size='10' maxlength='200'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵�����Ч������</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_10' id='RCM_Menu_10'>"
    Response.Write "        <option value='-1' "
    If RCM_Menu_10 = "-1" Then Response.Write " selected"
    Response.Write "        >����Ч</option>"
    Response.Write "        <option value='0' "
    If RCM_Menu_10 = "0" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_10 = "1" Then Response.Write " selected"
    Response.Write "        >������ɢ</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_10 = "2" Then Response.Write " selected"
    Response.Write "        >Բ������</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_10 = "3" Then Response.Write " selected"
    Response.Write "        >Բ����ɢ</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_10 = "4" Then Response.Write " selected"
    Response.Write "        >����Ч��</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_10 = "5" Then Response.Write " selected"
    Response.Write "        >����Ч��</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_10 = "6" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='7' "
    If RCM_Menu_10 = "7" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='8' "
    If RCM_Menu_10 = "8" Then Response.Write " selected"
    Response.Write "        >���Ұ�Ҷ</option>"
    Response.Write "        <option value='9' "
    If RCM_Menu_10 = "9" Then Response.Write " selected"
    Response.Write "        >���°�Ҷ</option>"
    Response.Write "        <option value='10' "
    If RCM_Menu_10 = "10" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='11' "
    If RCM_Menu_10 = "11" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='12' "
    If RCM_Menu_10 = "12" Then Response.Write " selected"
    Response.Write "        >ģ��Ч��</option>"
    Response.Write "        <option value='13' "
    If RCM_Menu_10 = "13" Then Response.Write " selected"
    Response.Write "        >���ҹ���</option>"
    Response.Write "        <option value='14' "
    If RCM_Menu_10 = "14" Then Response.Write " selected"
    Response.Write "        >���ҿ���</option>"
    Response.Write "        <option value='15' "
    If RCM_Menu_10 = "15" Then Response.Write " selected"
    Response.Write "        >���¹���</option>"
    Response.Write "        <option value='16' "
    If RCM_Menu_10 = "16" Then Response.Write " selected"
    Response.Write "        >���¿���</option>"
    Response.Write "        <option value='17' "
    If RCM_Menu_10 = "17" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='18' "
    If RCM_Menu_10 = "18" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='19' "
    If RCM_Menu_10 = "19" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='20' "
    If RCM_Menu_10 = "20" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='21' "
    If RCM_Menu_10 = "21" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='22' "
    If RCM_Menu_10 = "22" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='23' "
    If RCM_Menu_10 = "23" Then Response.Write " selected"
    Response.Write "        >�����Ч</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�����Ч������</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_12' id='RCM_Menu_12'>"
    Response.Write "        <option value='-1' "
    If RCM_Menu_12 = "-1" Then Response.Write " selected"
    Response.Write "        >����Ч</option>"
    Response.Write "        <option value='0' "
    If RCM_Menu_12 = "0" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_12 = "1" Then Response.Write " selected"
    Response.Write "        >������ɢ</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_12 = "2" Then Response.Write " selected"
    Response.Write "        >Բ������</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_12 = "3" Then Response.Write " selected"
    Response.Write "        >Բ����ɢ</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_12 = "4" Then Response.Write " selected"
    Response.Write "        >����Ч��</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_12 = "5" Then Response.Write " selected"
    Response.Write "        >����Ч��</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_12 = "6" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='7' "
    If RCM_Menu_12 = "7" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='8' "
    If RCM_Menu_12 = "8" Then Response.Write " selected"
    Response.Write "        >���Ұ�Ҷ</option>"
    Response.Write "        <option value='9' "
    If RCM_Menu_12 = "9" Then Response.Write " selected"
    Response.Write "        >���°�Ҷ</option>"
    Response.Write "        <option value='10' "
    If RCM_Menu_12 = "10" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='11' "
    If RCM_Menu_12 = "11" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='12' "
    If RCM_Menu_12 = "12" Then Response.Write " selected"
    Response.Write "        >ģ��Ч��</option>"
    Response.Write "        <option value='13' "
    If RCM_Menu_12 = "13" Then Response.Write " selected"
    Response.Write "        >���ҹ���</option>"
    Response.Write "        <option value='14' "
    If RCM_Menu_12 = "14" Then Response.Write " selected"
    Response.Write "        >���ҿ���</option>"
    Response.Write "        <option value='15' "
    If RCM_Menu_12 = "15" Then Response.Write " selected"
    Response.Write "        >���¹���</option>"
    Response.Write "        <option value='16' "
    If RCM_Menu_12 = "16" Then Response.Write " selected"
    Response.Write "        >���¿���</option>"
    Response.Write "        <option value='17' "
    If RCM_Menu_12 = "17" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='18' "
    If RCM_Menu_12 = "18" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='19' "
    If RCM_Menu_12 = "19" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='20' "
    If RCM_Menu_12 = "20" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='21' "
    If RCM_Menu_12 = "21" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='22' "
    If RCM_Menu_12 = "22" Then Response.Write " selected"
    Response.Write "        >��������</option>"
    Response.Write "        <option value='23' "
    If RCM_Menu_12 = "23" Then Response.Write " selected"
    Response.Write "        >�����Ч</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�����Ч���ٶȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_13' type='text' id='RCM_Menu_13' value='" & RCM_Menu_13 & "' size='10' maxlength='10' title='�ٶ�ֵ��10-100'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵���ӰЧ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_14' id='RCM_Menu_14'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_14 = "0" Then Response.Write " selected"
    Response.Write "        >����Ӱ</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_14 = "1" Then Response.Write " selected"
    Response.Write "        >����Ӱ</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_14 = "2" Then Response.Write " selected"
    Response.Write "        >������Ӱ</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵���Ӱ��ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_15' type='text' id='RCM_Menu_15' value='" & RCM_Menu_15 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵���Ӱ��ɫ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_16' type='text' id='RCM_Menu_16' value='" & RCM_Menu_16 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵�������ɫ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_17' type='text' id='RCM_Menu_17' value='" & RCM_Menu_17 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�����ͼƬ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_18' type='text' id='RCM_Menu_18' value='" & RCM_Menu_18 & "' size='10' maxlength='200' title='ֻ�е��˵������ɫ��Ϊ͸��ɫ��transparent ʱ����Ч'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>����ͼƬƽ��ģʽ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_19' id='RCM_Menu_19'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_19 = "0" Then Response.Write " selected"
    Response.Write "        >��ƽ��</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_19 = "1" Then Response.Write " selected"
    Response.Write "        >����ƽ��</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_19 = "2" Then Response.Write " selected"
    Response.Write "        >����ƽ��</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_19 = "3" Then Response.Write " selected"
    Response.Write "        >��ȫƽ��</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵��߿����ͣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_20' id='RCM_Menu_20'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_20 = "0" Then Response.Write " selected"
    Response.Write "        >�ޱ߿�</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_20 = "1" Then Response.Write " selected"
    Response.Write "        >��ʵ��</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_20 = "2" Then Response.Write " selected"
    Response.Write "        >˫ʵ��</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_20 = "5" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_20 = "6" Then Response.Write " selected"
    Response.Write "        >͹��</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵��߿��ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_21' type='text' id='RCM_Menu_21' value='" & RCM_Menu_21 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵��߿���ɫ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_22' type='text' id='RCM_Menu_22' value='" & RCM_Menu_22 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"

    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='6'><strong>�˵����������</strong></td>"
    Response.Write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>�˵������ͣ�</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <select name='RCM_Item_1' id='RCM_Item_1'>"
'    response.write "        <option value='0' "
'   if RCM_Menu_1="0" then response.write " selected"
'    response.write "        >�ı�</option>"
'    response.write "        <option value='1' "
'   if RCM_Menu_1="1" then response.write " selected"
'    response.write "        >HTML</option>"
'    response.write "        <option value='2' "
'   if RCM_Menu_1="2" then response.write " selected"
'    response.write "        >ͼƬ</option>"
'    response.write "      </select>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>�˵������ƣ�</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_2' type='text' id='RCM_Item_2' value='" & RCM_Item_2 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>ͼƬ�ļ���</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_3' type='text' id='RCM_Item_3' value='" & RCM_Item_3 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>���ָ�ڲ˵���ʱ��ͼƬ�ļ���</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_4' type='text' id='RCM_Item_4' value='" & RCM_Item_4 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>ͼƬ��ȣ�</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_5' type='text' id='RCM_Item_5' value='" & RCM_Item_5 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>ͼƬ�߶ȣ�</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_6' type='text' id='RCM_Item_6' value='" & RCM_Item_6 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>ͼƬ�߿�</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_7' type='text' id='RCM_Item_7' value='" & RCM_Item_7 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>���ӵ�ַ��</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_8' type='text' id='RCM_Item_8' value='" & RCM_Item_8 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>����Ŀ�꣺</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_9' type='text' id='RCM_Item_9' value='" & RCM_Item_9 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>����״̬����ʾ��</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_10' type='text' id='RCM_Item_10' value='" & RCM_Item_10 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>���ӵ�ַ��ʾ��Ϣ��</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_11' type='text' id='RCM_Item_11' value='" & RCM_Item_11 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong></strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      "
'    response.write "    </td>"
'    response.write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵�����ͼƬ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_12' type='text' id='RCM_Item_12' value='" & RCM_Item_12 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�����ͼƬ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_13' type='text' id='RCM_Item_13' value='" & RCM_Item_13 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>��ͼƬ��ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_14' type='text' id='RCM_Item_14' value='" & RCM_Item_14 & "' size='10' maxlength='10' title='0Ϊͼ��ԭʼ���'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>��ͼƬ�߶ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_15' type='text' id='RCM_Item_15' value='" & RCM_Item_15 & "' size='10' maxlength='10' title='0Ϊͼ��ԭʼ�߶�'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>��ͼƬ�߿��С��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_16' type='text' id='RCM_Item_16' value='" & RCM_Item_16 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵�����ͼƬ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_17' type='text' id='RCM_Item_17' value='" & RCM_Item_17 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵�����ͼƬ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_18' type='text' id='RCM_Item_18' value='" & RCM_Item_18 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>��ͼƬ��ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_19' type='text' id='RCM_Item_19' value='" & RCM_Item_19 & "' size='10' maxlength='10' title='0Ϊͼ��ԭʼ���'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>��ͼƬ�߶ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_20' type='text' id='RCM_Item_20' value='" & RCM_Item_20 & "' size='10' maxlength='10' title='0Ϊͼ��ԭʼ�߶�'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>��ͼƬ�߿��С��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_21' type='text' id='RCM_Item_21' value='" & RCM_Item_21 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>����ˮƽ���뷽ʽ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_22' id='RCM_Item_22'>"
    Response.Write "        <option value='0' "
    If RCM_Item_22 = "0" Then Response.Write " selected"
    Response.Write "        >�����</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_22 = "1" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_22 = "2" Then Response.Write " selected"
    Response.Write "        >�Ҷ���</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>���ִ�ֱ���뷽ʽ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_23' id='RCM_Item_23'>"
    Response.Write "        <option value='0' "
    If RCM_Item_23 = "0" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_23 = "1" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_23 = "2" Then Response.Write " selected"
    Response.Write "        >�ײ�</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵������ɫ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_24' type='text' id='RCM_Item_24' value='" & RCM_Item_24 & "' size='10' maxlength='10' title='͸��ɫ��transparent'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>������ɫ�Ƿ���ʾ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_25' id='RCM_Item_25'>"
    Response.Write "        <option value='0' "
    If RCM_Item_25 = "0" Then Response.Write " selected"
    Response.Write "        >��ʾ</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_25 = "1" Then Response.Write " selected"
    Response.Write "        >����ʾ</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵������ɫ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_26' type='text' id='RCM_Item_26' value='" & RCM_Item_26 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>������ɫ�Ƿ���ʾ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_27' id='RCM_Item_27'>"
    Response.Write "        <option value='0' "
    If RCM_Item_27 = "0" Then Response.Write " selected"
    Response.Write "        >��ʾ</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_27 = "1" Then Response.Write " selected"
    Response.Write "        >����ʾ</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵����ͼƬ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_28' type='text' id='RCM_Item_28' value='" & RCM_Item_28 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵����ͼƬ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_29' type='text' id='RCM_Item_29' value='" & RCM_Item_29 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>����ͼƬƽ��ģʽ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_30' id='RCM_Item_30'>"
    Response.Write "        <option value='0' "
    If RCM_Item_30 = "0" Then Response.Write " selected"
    Response.Write "        >��ƽ��</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_30 = "1" Then Response.Write " selected"
    Response.Write "        >����ƽ��</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_30 = "2" Then Response.Write " selected"
    Response.Write "        >����ƽ��</option>"
    Response.Write "        <option value='3' "
    If RCM_Item_30 = "3" Then Response.Write " selected"
    Response.Write "        >��ȫƽ��</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
'    response.write "    <td width='130' height='25'><strong>����ͼƬƽ��ģʽ����</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <select name='RCM_Item_31' id='RCM_Item_31'>"
'    response.write "        <option value='0' "
'   if RCM_Menu_1="0" then response.write " selected"
'    response.write "        >��ƽ��</option>"
'    response.write "        <option value='1' "
'   if RCM_Menu_1="1" then response.write " selected"
'    response.write "        >����ƽ��</option>"
'    response.write "        <option value='2' "
'   if RCM_Menu_1="2" then response.write " selected"
'    response.write "        >����ƽ��</option>"
'    response.write "        <option value='3' "
'   if RCM_Menu_1="3" then response.write " selected"
'    response.write "        >��ȫƽ��</option>"
'    response.write "      </select>"
'    response.write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵���߿����ͣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_32' id='RCM_Item_32'>"
    Response.Write "        <option value='0' "
    If RCM_Item_32 = "0" Then Response.Write " selected"
    Response.Write "        >�ޱ߿�</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_32 = "1" Then Response.Write " selected"
    Response.Write "        >��ʵ��</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_32 = "2" Then Response.Write " selected"
    Response.Write "        >˫ʵ��</option>"
    Response.Write "        <option value='5' "
    If RCM_Item_32 = "5" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value='6' "
    If RCM_Item_32 = "6" Then Response.Write " selected"
    Response.Write "        >͹��</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵���߿��ȣ�</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_33' type='text' id='RCM_Item_33' value='" & RCM_Item_33 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵���߿���ɫ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_34' type='text' id='RCM_Item_34' value='" & RCM_Item_34 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵���߿���ɫ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_35' type='text' id='RCM_Item_35' value='" & RCM_Item_35 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵���������ɫ��</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_36' type='text' id='RCM_Item_36' value='" & RCM_Item_36 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>�˵���������ɫ����</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_37' type='text' id='RCM_Item_37' value='" & RCM_Item_37 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵����������壺</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='FontName_RCM_Item_38' id='FontName_RCM_Item_38'>"
    Response.Write "        <option value='����' "
    If FontName_RCM_Item_38 = "����" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����"" "
    If FontName_RCM_Item_38 = "����" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����_GB2312"" "
    If FontName_RCM_Item_38 = "����_GB2312" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����_GB2312"" "
    If FontName_RCM_Item_38 = "����_GB2312" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����"" "
    If FontName_RCM_Item_38 = "����" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""��Բ"" "
    If FontName_RCM_Item_38 = "��Բ" Then Response.Write " selected"
    Response.Write "        >��Բ</option>"
    Response.Write "        <option value=""Arial"" "
    If FontName_RCM_Item_38 = "Arial" Then Response.Write " selected"
    Response.Write "        >Arial</option>"
    Response.Write "        <option value=""Arial Black"" "
    If FontName_RCM_Item_38 = "Arial Black" Then Response.Write " selected"
    Response.Write "        >Arial Black</option>"
    Response.Write "        <option value=""Arial Narrow"" "
    If FontName_RCM_Item_38 = "Arial Narrow" Then Response.Write " selected"
    Response.Write "        >Arial Narrow</option>"
    Response.Write "        <option value=""Brush ScriptMT"" "
    If FontName_RCM_Item_38 = "Brush ScriptMT" Then Response.Write " selected"
    Response.Write "        >Brush Script MT</option>"
    Response.Write "        <option value=""Century Gothic"" "
    If FontName_RCM_Item_38 = "Century Gothic" Then Response.Write " selected"
    Response.Write "        >Century Gothic</option>"
    Response.Write "        <option value=""Comic Sans MS"" "
    If FontName_RCM_Item_38 = "Comic Sans MS" Then Response.Write " selected"
    Response.Write "        >Comic Sans MS</option>"
    Response.Write "        <option value=""Courier"" "
    If FontName_RCM_Item_38 = "Courier" Then Response.Write " selected"
    Response.Write "        >Courier</option>"
    Response.Write "        <option value=""Courier New"" "
    If FontName_RCM_Item_38 = "Courier New" Then Response.Write " selected"
    Response.Write "        >Courier New</option>"
    Response.Write "        <option value=""MS Sans Serif"" "
    If FontName_RCM_Item_38 = "MS Sans Serif" Then Response.Write " selected"
    Response.Write "        >MS Sans Serif</option>"
    Response.Write "        <option value=""Script"" "
    If FontName_RCM_Item_38 = "Script" Then Response.Write " selected"
    Response.Write "        >Script</option>"
    Response.Write "        <option value=""System"" "
    If FontName_RCM_Item_38 = "System" Then Response.Write " selected"
    Response.Write "        >System</option>"
    Response.Write "        <option value=""Times New Roman"" "
    If FontName_RCM_Item_38 = "Times New Roman" Then Response.Write " selected"
    Response.Write "        >Times New Roman</option>"
    Response.Write "        <option value=""Verdana"" "
    If FontName_RCM_Item_38 = "Verdana" Then Response.Write " selected"
    Response.Write "        >Verdana</option>"
    Response.Write "        <option value=""WideLatin"" "
    If FontName_RCM_Item_38 = "WideLatin" Then Response.Write " selected"
    Response.Write "        >Wide Latin</option>"
    Response.Write "        <option value=""Wingdings"" "
    If FontName_RCM_Item_38 = "Wingdings" Then Response.Write " selected"
    Response.Write "        >Wingdings</option>"
    Response.Write "      </select>"
    Response.Write "      <select name = 'FontSize_RCM_Item_38' id='FontSize_RCM_Item_38'>"
    Response.Write "        <option value=""9pt"" "
    If FontSize_RCM_Item_38 = "9pt" Then Response.Write " selected"
    Response.Write "        >9pt</option>"
    Response.Write "        <option value=""10pt"" "
    If FontSize_RCM_Item_38 = "10pt" Then Response.Write " selected"
    Response.Write "        >10pt</option>"
    Response.Write "        <option value=""12pt"" "
    If FontSize_RCM_Item_38 = "12pt" Then Response.Write " selected"
    Response.Write "        >12pt</option>"
    Response.Write "        <option value=""14pt"" "
    If FontSize_RCM_Item_38 = "14pt" Then Response.Write " selected"
    Response.Write "        >14pt</option>"
    Response.Write "        <option value=""16pt"" "
    If FontSize_RCM_Item_38 = "16pt" Then Response.Write " selected"
    Response.Write "        >16pt</option>"
    Response.Write "        <option value=""18pt"" "
    If FontSize_RCM_Item_38 = "18pt" Then Response.Write " selected"
    Response.Write "        >18pt</option>"
    Response.Write "        <option value=""24pt"" "
    If FontSize_RCM_Item_38 = "24pt" Then Response.Write " selected"
    Response.Write "        >24pt</option>"
    Response.Write "        <option value=""36pt"" "
    If FontSize_RCM_Item_38 = "36pt" Then Response.Write " selected"
    Response.Write "        >36pt</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>�˵��������������</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='FontName_RCM_Item_39' id='FontName_RCM_Item_39'>"
    Response.Write "        <option value='����' "
    If FontName_RCM_Item_39 = "����" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����"" "
    If FontName_RCM_Item_39 = "����" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����_GB2312"" "
    If FontName_RCM_Item_39 = "����_GB2312" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����_GB2312"" "
    If FontName_RCM_Item_39 = "����_GB2312" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""����"" "
    If FontName_RCM_Item_39 = "����" Then Response.Write " selected"
    Response.Write "        >����</option>"
    Response.Write "        <option value=""��Բ"" "
    If FontName_RCM_Item_39 = "��Բ" Then Response.Write " selected"
    Response.Write "        >��Բ</option>"
    Response.Write "        <option value=""Arial"" "
    If FontName_RCM_Item_39 = "Arial" Then Response.Write " selected"
    Response.Write "        >Arial</option>"
    Response.Write "        <option value=""Arial Black"" "
    If FontName_RCM_Item_39 = "Arial Black" Then Response.Write " selected"
    Response.Write "        >Arial Black</option>"
    Response.Write "        <option value=""Arial Narrow"" "
    If FontName_RCM_Item_39 = "Arial Narrow" Then Response.Write " selected"
    Response.Write "        >Arial Narrow</option>"
    Response.Write "        <option value=""Brush ScriptMT"" "
    If FontName_RCM_Item_39 = "Brush ScriptMT" Then Response.Write " selected"
    Response.Write "        >Brush Script MT</option>"
    Response.Write "        <option value=""Century Gothic"" "
    If FontName_RCM_Item_39 = "Century Gothic" Then Response.Write " selected"
    Response.Write "        >Century Gothic</option>"
    Response.Write "        <option value=""Comic Sans MS"" "
    If FontName_RCM_Item_39 = "Comic Sans MS" Then Response.Write " selected"
    Response.Write "        >Comic Sans MS</option>"
    Response.Write "        <option value=""Courier"" "
    If FontName_RCM_Item_39 = "Courier" Then Response.Write " selected"
    Response.Write "        >Courier</option>"
    Response.Write "        <option value=""Courier New"" "
    If FontName_RCM_Item_39 = "Courier New" Then Response.Write " selected"
    Response.Write "        >Courier New</option>"
    Response.Write "        <option value=""MS Sans Serif"" "
    If FontName_RCM_Item_39 = "MS Sans Serif" Then Response.Write " selected"
    Response.Write "        >MS Sans Serif</option>"
    Response.Write "        <option value=""Script"" "
    If FontName_RCM_Item_39 = "Script" Then Response.Write " selected"
    Response.Write "        >Script</option>"
    Response.Write "        <option value=""System"" "
    If FontName_RCM_Item_39 = "System" Then Response.Write " selected"
    Response.Write "        >System</option>"
    Response.Write "        <option value=""Times New Roman"" "
    If FontName_RCM_Item_39 = "Times New Roman" Then Response.Write " selected"
    Response.Write "        >Times New Roman</option>"
    Response.Write "        <option value=""Verdana"" "
    If FontName_RCM_Item_39 = "Verdana" Then Response.Write " selected"
    Response.Write "        >Verdana</option>"
    Response.Write "        <option value=""WideLatin"" "
    If FontName_RCM_Item_39 = "WideLatin" Then Response.Write " selected"
    Response.Write "        >Wide Latin</option>"
    Response.Write "        <option value=""Wingdings"" "
    If FontName_RCM_Item_39 = "Wingdings" Then Response.Write " selected"
    Response.Write "        >Wingdings</option>"
    Response.Write "      </select>"
    Response.Write "      <select name = 'FontSize_RCM_Item_39' id='FontSize_RCM_Item_39'>"
    Response.Write "        <option value=""9pt"" "
    If FontSize_RCM_Item_39 = "9pt" Then Response.Write " selected"
    Response.Write "        >9pt</option>"
    Response.Write "        <option value=""10pt"" "
    If FontSize_RCM_Item_39 = "10pt" Then Response.Write " selected"
    Response.Write "        >10pt</option>"
    Response.Write "        <option value=""12pt"" "
    If FontSize_RCM_Item_39 = "12pt" Then Response.Write " selected"
    Response.Write "        >12pt</option>"
    Response.Write "        <option value=""14pt"" "
    If FontSize_RCM_Item_39 = "14pt" Then Response.Write " selected"
    Response.Write "        >14pt</option>"
    Response.Write "        <option value=""16pt"" "
    If FontSize_RCM_Item_39 = "16pt" Then Response.Write " selected"
    Response.Write "        >16pt</option>"
    Response.Write "        <option value=""18pt"" "
    If FontSize_RCM_Item_39 = "18pt" Then Response.Write " selected"
    Response.Write "        >18pt</option>"
    Response.Write "        <option value=""24pt"" "
    If FontSize_RCM_Item_39 = "24pt" Then Response.Write " selected"
    Response.Write "        >24pt</option>"
    Response.Write "        <option value=""36pt"" "
    If FontSize_RCM_Item_39 = "36pt" Then Response.Write " selected"
    Response.Write "        >36pt</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' colspan='6' align='center'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveConfig'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "      <input name='cmdSave' type='submit' id='cmdSave' value=' �������� ' "
    If ObjInstalled_FSO = False Then Response.Write " disabled"
    Response.Write "      >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub SaveConfig()
    If ObjInstalled_FSO = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��ķ�������֧�� FSO(Scripting.FileSystemObject)! </li>"
        Exit Sub
    End If
    Set hf = fso.CreateTextFile(Server.MapPath(InstallDir & AdminDir & "/RootClass_Menu_Config.asp"), True)

    hf.Write "<" & "%" & vbCrLf
    hf.Write "'�˵���ʾ��������" & vbCrLf
    hf.Write "Const RCM_Menu_1=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_1"))) & Chr(34) & "      '�˵�������ʽ 1����  2����  3����  4����" & vbCrLf
    hf.Write "Const RCM_Menu_2=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_2"))) & Chr(34) & "      '�˵���������ƫ����" & vbCrLf
    hf.Write "Const RCM_Menu_3=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_3"))) & Chr(34) & "      '�˵���������ƫ����" & vbCrLf
    hf.Write "Const RCM_Menu_4=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_4"))) & Chr(34) & "      '�˵���߾�" & vbCrLf
    hf.Write "Const RCM_Menu_5=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_5"))) & Chr(34) & "      '�˵�����" & vbCrLf
    hf.Write "Const RCM_Menu_6=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_6"))) & Chr(34) & "      '�˵�����߾�" & vbCrLf
    hf.Write "Const RCM_Menu_7=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_7"))) & Chr(34) & "      '�˵����ұ߾�" & vbCrLf
    hf.Write "Const RCM_Menu_8=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_8"))) & Chr(34) & "      '�˵�͸����         0-100 ��ȫ͸��-��ȫ��͸��" & vbCrLf
    hf.Write "Const RCM_Menu_9=" & Chr(34) & FilterString(Trim(request("RCM_Menu_9"))) & Chr(34) & "      '������Ч" & vbCrLf
    hf.Write "Const RCM_Menu_10=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_10"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵�����Ч��" & vbCrLf
    hf.Write "Const RCM_Menu_11=" & Chr(34) & FilterString(Trim(request("RCM_Menu_11"))) & Chr(34) & "        '������Ч" & vbCrLf
    hf.Write "Const RCM_Menu_12=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_12"))) & Chr(34) & "        '����Ƴ��˵���ʱ���˵�����Ч��" & vbCrLf
    hf.Write "Const RCM_Menu_13=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_13"))) & Chr(34) & "        '�˵�����Ч���ٶ�  10-100" & vbCrLf
    hf.Write "Const RCM_Menu_14=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_14"))) & Chr(34) & "        '�����˵���ӰЧ�� 0��none  1��simple  2��complex" & vbCrLf
    hf.Write "Const RCM_Menu_15=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_15"))) & Chr(34) & "        '�����˵���Ӱ���" & vbCrLf
    hf.Write "Const RCM_Menu_16=" & Chr(34) & FilterString(Trim(request("RCM_Menu_16"))) & Chr(34) & "        '�����˵���Ӱ��ɫ" & vbCrLf
    hf.Write "Const RCM_Menu_17=" & Chr(34) & FilterString(Trim(request("RCM_Menu_17"))) & Chr(34) & "        '�����˵�������ɫ" & vbCrLf
    hf.Write "Const RCM_Menu_18=" & Chr(34) & FilterString(Trim(request("RCM_Menu_18"))) & Chr(34) & "        '�����˵�����ͼƬ��ֻ�е��˵������ɫ��Ϊ͸��ɫ��transparent ʱ����Ч" & vbCrLf
    hf.Write "Const RCM_Menu_19=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_19"))) & Chr(34) & "        '�����˵�����ͼƬƽ��ģʽ�� 0����ƽ��  1������ƽ��  2������ƽ��  3����ȫƽ��" & vbCrLf
    hf.Write "Const RCM_Menu_20=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_20"))) & Chr(34) & "        '�����˵��߿����� 0���ޱ߿�  1����ʵ��  2��˫ʵ��  5������  6��͹��" & vbCrLf
    hf.Write "Const RCM_Menu_21=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_21"))) & Chr(34) & "        '�����˵��߿���" & vbCrLf
    hf.Write "Const RCM_Menu_22=" & Chr(34) & FilterString(Trim(request("RCM_Menu_22"))) & Chr(34) & "        '�����˵��߿���ɫ" & vbCrLf
    hf.Write "Const RCM_Menu_23=" & Chr(34) & "#ffffff" & Chr(34) & "" & vbCrLf
    hf.Write "" & vbCrLf
    hf.Write "'�˵����������" & vbCrLf
    hf.Write "Const RCM_Item_1=" & Chr(34) & "0" & Chr(34) & "      '�˵�������  0--Txt  1--Html  2--Image" & vbCrLf
    hf.Write "Const RCM_Item_2=" & Chr(34) & "" & Chr(34) & "       '�˵�������" & vbCrLf
    hf.Write "Const RCM_Item_3=" & Chr(34) & "" & Chr(34) & "       '�˵���ΪImage��ͼƬ�ļ�" & vbCrLf
    hf.Write "Const RCM_Item_4=" & Chr(34) & "" & Chr(34) & "       '�˵���ΪImage�����ָ�ڲ˵���ʱ��ͼƬ�ļ���" & vbCrLf
    hf.Write "Const RCM_Item_5=" & Chr(34) & "-1" & Chr(34) & "     '�˵���ΪImage��ͼƬ���" & vbCrLf
    hf.Write "Const RCM_Item_6=" & Chr(34) & "-1" & Chr(34) & "     '�˵���ΪImage��ͼƬ�߶�" & vbCrLf
    hf.Write "Const RCM_Item_7=" & Chr(34) & "0" & Chr(34) & "      '�˵���ΪImage��ͼƬ�߿�" & vbCrLf
    hf.Write "Const RCM_Item_8=" & Chr(34) & "" & Chr(34) & "       '�˵������ӵ�ַ" & vbCrLf
    hf.Write "Const RCM_Item_9=" & Chr(34) & "" & Chr(34) & "       '�˵�������Ŀ�� �磺_self  _blank" & vbCrLf
    hf.Write "Const RCM_Item_10=" & Chr(34) & "" & Chr(34) & "      '�˵�������״̬����ʾ" & vbCrLf
    hf.Write "Const RCM_Item_11=" & Chr(34) & "" & Chr(34) & "      '�˵������ӵ�ַ��ʾ��Ϣ" & vbCrLf
    hf.Write "Const RCM_Item_12=" & Chr(34) & FilterString(Trim(request("RCM_Item_12"))) & Chr(34) & "        '�˵�����ͼƬ" & vbCrLf
    hf.Write "Const RCM_Item_13=" & Chr(34) & FilterString(Trim(request("RCM_Item_13"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵�����ͼƬ" & vbCrLf
    hf.Write "Const RCM_Item_14=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_14"))) & Chr(34) & "        '�˵�����ͼƬ��ȣ�0Ϊͼ���ļ�ԭʼֵ" & vbCrLf
    hf.Write "Const RCM_Item_15=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_15"))) & Chr(34) & "        '�˵�����ͼƬ�߶ȣ�0Ϊͼ���ļ�ԭʼֵ" & vbCrLf
    hf.Write "Const RCM_Item_16=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_16"))) & Chr(34) & "        '�˵�����ͼƬ�߿��С" & vbCrLf
    hf.Write "Const RCM_Item_17=" & Chr(34) & FilterString(Trim(request("RCM_Item_17"))) & Chr(34) & "        '�˵�����ͼƬ���磺arrow_r.gif" & vbCrLf
    hf.Write "Const RCM_Item_18=" & Chr(34) & FilterString(Trim(request("RCM_Item_18"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵�����ͼƬ���磺arrow_w.gif" & vbCrLf
    hf.Write "Const RCM_Item_19=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_19"))) & Chr(34) & "        '�˵�����ͼƬ��ȣ�0Ϊͼ���ļ�ԭʼֵ" & vbCrLf
    hf.Write "Const RCM_Item_20=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_20"))) & Chr(34) & "        '�˵�����ͼƬ�߶ȣ�0Ϊͼ���ļ�ԭʼֵ" & vbCrLf
    hf.Write "Const RCM_Item_21=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_21"))) & Chr(34) & "        '�˵�����ͼƬ�߿��С" & vbCrLf
    hf.Write "Const RCM_Item_22=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_22"))) & Chr(34) & "        '�˵�������ˮƽ���뷽ʽ  0�������  1������  2���Ҷ���" & vbCrLf
    hf.Write "Const RCM_Item_23=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_23"))) & Chr(34) & "        '�˵������ִ�ֱ���뷽ʽ  0������  1������  2���ײ�" & vbCrLf
    hf.Write "Const RCM_Item_24=" & Chr(34) & FilterString(Trim(request("RCM_Item_24"))) & Chr(34) & "        '�˵������ɫ  ͸��ɫ��'transparent'" & vbCrLf
    hf.Write "Const RCM_Item_25=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_25"))) & Chr(34) & "        '�˵������ɫ�Ƿ���ʾ  0����ʾ  ����������ʾ" & vbCrLf
    hf.Write "Const RCM_Item_26=" & Chr(34) & FilterString(Trim(request("RCM_Item_26"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵������ɫ" & vbCrLf
    hf.Write "Const RCM_Item_27=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_27"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵������ɫ�Ƿ���ʾ��  0����ʾ  ����������ʾ" & vbCrLf
    hf.Write "Const RCM_Item_28=" & Chr(34) & FilterString(Trim(request("RCM_Item_28"))) & Chr(34) & "        '�˵����ͼƬ" & vbCrLf
    hf.Write "Const RCM_Item_29=" & Chr(34) & FilterString(Trim(request("RCM_Item_29"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵����ͼƬ" & vbCrLf
    hf.Write "Const RCM_Item_30=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_30"))) & Chr(34) & "        '�˵����ͼƬƽ��ģʽ�� 0����ƽ��  1������ƽ��  2������ƽ��  3����ȫƽ��" & vbCrLf
    hf.Write "Const RCM_Item_31=" & Chr(34) & "3" & Chr(34) & "     '���ָ�ڲ˵���ʱ���˵����ͼƬƽ��ģʽ��0-3" & vbCrLf
    hf.Write "Const RCM_Item_32=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_32"))) & Chr(34) & "        '�˵���߿����� 0���ޱ߿�  1����ʵ��  2��˫ʵ��  5������  6��͹��" & vbCrLf
    hf.Write "Const RCM_Item_33=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_33"))) & Chr(34) & "        '�˵���߿���" & vbCrLf
    hf.Write "Const RCM_Item_34=" & Chr(34) & FilterString(Trim(request("RCM_Item_34"))) & Chr(34) & "        '�˵���߿���ɫ" & vbCrLf
    hf.Write "Const RCM_Item_35=" & Chr(34) & FilterString(Trim(request("RCM_Item_35"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵���߿���ɫ" & vbCrLf
    hf.Write "Const RCM_Item_36=" & Chr(34) & FilterString(Trim(request("RCM_Item_36"))) & Chr(34) & "        '�˵���������ɫ" & vbCrLf
    hf.Write "Const RCM_Item_37=" & Chr(34) & FilterString(Trim(request("RCM_Item_37"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ���˵���������ɫ" & vbCrLf
    hf.Write "Const FontSize_RCM_Item_38=" & Chr(34) & FilterString(Trim(request("FontSize_RCM_Item_38"))) & Chr(34) & "        '�˵������ִ�С" & vbCrLf
    hf.Write "Const FontName_RCM_Item_38=" & Chr(34) & FilterString(Trim(request("FontName_RCM_Item_38"))) & Chr(34) & "        '�˵�����������" & vbCrLf
    hf.Write "Const FontSize_RCM_Item_39=" & Chr(34) & FilterString(Trim(request("FontSize_RCM_Item_39"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ,�˵������ִ�С" & vbCrLf
    hf.Write "Const FontName_RCM_Item_39=" & Chr(34) & FilterString(Trim(request("FontName_RCM_Item_39"))) & Chr(34) & "        '���ָ�ڲ˵���ʱ,�˵�����������" & vbCrLf
    hf.Write "%" & ">"
    hf.Close
    Call WriteSuccessMsg("������Ŀ�˵��������óɹ���", ComeUrl)
End Sub

Sub ShowCreate_RootClass_Menu()
    Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong> �� �� �� �� �� Ŀ �� �� </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150'>"
    Response.Write "<form name='myform' method='post' action='Admin_RootClass_Menu.asp'>"
    Response.Write "<p align='center'>�˲��������ݶ�����Ŀ�˵��������������õĲ��������Զ���Ĳ˵���</p>"
    Response.Write "<p align='center'><input name='Action' type='hidden' id='Action' value='Create'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='submit' name='Submit' value=' ���ɶ�����Ŀ�˵� '></p>"
    Response.Write "</form>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub Create_RootClass_Menu()
    strTopMenu = GetRootClass_Menu()
    If Not fso.FolderExists(Server.MapPath(InstallDir & ChannelDir & "/js")) Then
        fso.CreateFolder Server.MapPath(InstallDir & ChannelDir & "/js")
    End If
    Set hf = fso.CreateTextFile(Server.MapPath(InstallDir & ChannelDir & "/js/ShowClass_Menu.js"), True)
    hf.Write strTopMenu
    hf.Close
    Call WriteSuccessMsg("������Ŀ�˵����ɳɹ���", ComeUrl)
End Sub

'=================================================
'��������GetRootClass_Menu
'��  �ã��õ���Ŀ�޼������˵�Ч����HTML����
'��  ������
'����ֵ����Ŀ�޼������˵�Ч����HTML����
'=================================================
Function GetRootClass_Menu()
    Dim Class_MenuTitle, strJS
    pNum = 1
    pNum2 = 0
    strJS = stm_bm() & vbCrLf
    strJS = strJS & stm_bp_h() & vbCrLf
    strJS = strJS & stm_ai() & vbCrLf
    If UseCreateHTML > 0 Then
        strJS = strJS & stm_aix("p0i1", "p0i0", ChannelName & "��ҳ", ChannelUrl & "/Index" & FileExt_List, "_self", "", False) & vbCrLf
    Else
        strJS = strJS & stm_aix("p0i1", "p0i0", ChannelName & "��ҳ", ChannelUrl & "/Index.asp", "_self", "", False) & vbCrLf
    End If
    strJS = strJS & stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf

    Dim sqlRoot, rsRoot, j
    sqlRoot = "select * from PE_Class where ChannelID=" & ChannelID & " and Depth=0 and ShowOnTop=" & PE_True & " order by RootID"
    Set rsRoot = Server.CreateObject("ADODB.Recordset")
    rsRoot.open sqlRoot, Conn, 1, 1
    If Not (rsRoot.bof And rsRoot.EOF) Then
        j = 3
        Do While Not rsRoot.EOF
            If rsRoot("OpenType") = 0 Then
                OpenType_Class = "_self"
            Else
                OpenType_Class = "_blank"
            End If
            If Trim(rsRoot("Tips")) <> "" Then
                Class_MenuTitle = Replace(Replace(Replace(Replace(rsRoot("Tips"), "'", ""), """", ""), Chr(10), ""), Chr(13), "")
            Else
                Class_MenuTitle = ""
            End If
            If rsRoot("ClassType") = 1 Then
                If UseCreateHTML > 0 And rsRoot("ClassPurview") < 2 and UseCreateHTML<>2 Then
                    Select Case ListFileType
                    Case 0
                        strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot("ClassName"), ChannelUrl & rsRoot("ParentDir") & rsRoot("ClassDir") & "/Index" & FileExt_List, OpenType_Class, Class_MenuTitle, False) & vbCrLf
                    Case 1
                        strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot("ClassName"), ChannelUrl & "/List/List_" & rsRoot("ClassID") & FileExt_List, OpenType_Class, Class_MenuTitle, False) & vbCrLf
                    Case 2
                        strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot("ClassName"), ChannelUrl & "/List_" & rsRoot("ClassID") & FileExt_List, OpenType_Class, Class_MenuTitle, False) & vbCrLf
                    Case Else
                        strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot("ClassName"), ChannelUrl & "/ShowClass.asp?ClassID=" & rsRoot("ClassID"), OpenType_Class, Class_MenuTitle, False) & vbCrLf
                    End Select
                Else
                    strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot("ClassName"), ChannelUrl & "/ShowClass.asp?ClassID=" & rsRoot("ClassID"), OpenType_Class, Class_MenuTitle, False) & vbCrLf
                End If
                If rsRoot("Child") > 0 Then
                    strJS = strJS & GetClassMenu(rsRoot("ClassID"), 0)
                End If
            Else
                strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot("ClassName"), rsRoot("LinkUrl"), OpenType_Class, Class_MenuTitle, False) & vbCrLf
            End If
            strJS = strJS & stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
            j = j + 1
            rsRoot.movenext
            If (j - 2) Mod MaxPerLine = 0 And Not rsRoot.EOF Then
                strJS = strJS & "stm_em();" & vbCrLf
                strJS = strJS & stm_bm() & vbCrLf
                strJS = strJS & stm_bp_h() & vbCrLf
                strJS = strJS & stm_ai() & vbCrLf
            End If
        Loop
    End If
    rsRoot.Close
    Set rsRoot = Nothing
    strJS = strJS & "stm_em();" & vbCrLf

    GetRootClass_Menu = strJS
End Function

Function GetClassUrl(sParentDir, sClassDir, iClassID, iClassPurview)
    Dim strClassUrl
    If (UseCreateHTML = 1 Or UseCreateHTML = 3) And iClassPurview < 2 Then
        strClassUrl = ChannelUrl & GetListPath(StructureType, ListFileType, sParentDir, sClassDir) & GetListFileName(ListFileType, iClassID, 1, 1) & FileExt_List
    Else
        strClassUrl = ChannelUrl & "/ShowClass.asp?ClassID=" & iClassID
    End If
    GetClassUrl = strClassUrl
End Function

Function GetClassMenu(ID, ShowType)
    Dim sqlClass, rsClass, Sub_MenuTitle, k, strJS
    strJS = ""
    If pNum = 1 Then
        strJS = strJS & stm_bp_v("p" & pNum & "") & vbCrLf
    Else
        strJS = strJS & stm_bpx("p" & pNum & "", "p" & pNum2 & "", ShowType) & vbCrLf
    End If
    
    k = 0
    sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " and ParentID=" & ID & " order by OrderID asc"
    Set rsClass = Server.CreateObject("ADODB.Recordset")
    rsClass.open sqlClass, Conn, 1, 1
    Do While Not rsClass.EOF
        If rsClass("OpenType") = 0 Then
            OpenType_Class = "_self"
        Else
            OpenType_Class = "_blank"
        End If
        If Trim(rsClass("Tips")) <> "" Then
            Sub_MenuTitle = Replace(Replace(Replace(Replace(rsClass("Tips"), "'", ""), """", ""), Chr(10), ""), Chr(13), "")
        Else
            Sub_MenuTitle = ""
        End If
        If rsClass("ClassType") = 1 Then
            Dim strClassUrl
            strClassUrl = GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), rsClass("ClassPurview"))
            If rsClass("Child") > 0 Then
                If UseCreateHTML > 0 And rsClass("ClassPurview") < 2  and UseCreateHTML<>2 Then
                    Select Case ListFileType
                    Case 0
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), strClassUrl, OpenType_Class, Sub_MenuTitle, True) & vbCrLf
                    Case 1
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/List/List_" & rsClass("ClassID") & FileExt_List, OpenType_Class, Sub_MenuTitle, True) & vbCrLf
                    Case 2
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/List_" & rsClass("ClassID") & FileExt_List, OpenType_Class, Sub_MenuTitle, True) & vbCrLf
                    Case Else
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/ShowClass.asp?ClassID=" & rsClass("ClassID"), OpenType_Class, Sub_MenuTitle, True) & vbCrLf
                    End Select
                Else
                    strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/ShowClass.asp?ClassID=" & rsClass("ClassID"), OpenType_Class, Sub_MenuTitle, True) & vbCrLf
                End If
                pNum = pNum + 1
                pNum2 = pNum2 + 1
                strJS = strJS & GetClassMenu(rsClass("ClassID"), 1)
            Else
                If UseCreateHTML > 0 And rsClass("ClassPurview") < 2 Then
                    Select Case ListFileType
                    Case 0
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"),strClassUrl , OpenType_Class, Sub_MenuTitle, False) & vbCrLf
                    Case 1
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/List/List_" & rsClass("ClassID") & FileExt_List, OpenType_Class, Sub_MenuTitle, False) & vbCrLf
                    Case 2
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/List_" & rsClass("ClassID") & FileExt_List, OpenType_Class, Sub_MenuTitle, False) & vbCrLf
                    Case Else
                        strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/ShowClass.asp?ClassID=" & rsClass("ClassID"), OpenType_Class, Sub_MenuTitle, False) & vbCrLf
                    End Select
                Else
                    strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), ChannelUrl & "/ShowClass.asp?ClassID=" & rsClass("ClassID"), OpenType_Class, Sub_MenuTitle, False) & vbCrLf
                End If
            End If
        Else
            strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("ClassName"), rsClass("LinkUrl"), OpenType_Class, Sub_MenuTitle, False) & vbCrLf
        End If
        k = k + 1
        rsClass.movenext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    strJS = strJS & "stm_ep();" & vbCrLf

    GetClassMenu = strJS
End Function

Function stm_bm()
    stm_bm = "stm_bm(['uueoehr',400,'','" & strInstallDir & "images/blank.gif',0,'','',0,0,0,0,0,1,0,0]);"
End Function

Function stm_bp_h()
    stm_bp_h = "stm_bp('p0',[0,4,0,0,2,2,0,0," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & ",0,0,'#000000','transparent','',3,0,0,'#000000']);"
End Function

Function stm_bp_v(bpID)
    stm_bp_v = "stm_bp('" & bpID & "',[1," & RCM_Menu_1 & "," & RCM_Menu_2 & "," & RCM_Menu_3 & "," & RCM_Menu_4 & "," & RCM_Menu_5 & "," & RCM_Menu_6 & "," & RCM_Menu_7 & "," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & "," & RCM_Menu_14 & "," & RCM_Menu_15 & ",'" & RCM_Menu_16 & "','" & RCM_Menu_17 & "','" & RCM_Menu_18 & "'," & RCM_Menu_19 & "," & RCM_Menu_20 & "," & RCM_Menu_21 & ",'" & RCM_Menu_22 & "']);"
End Function

Function stm_bpx(bpOID, bpTID, bpType)
    If bpType = 0 Then
        stm_bpx = "stm_bpx('" & bpOID & "','" & bpTID & "',[1," & RCM_Menu_1 & "," & RCM_Menu_2 & "," & RCM_Menu_3 & "," & RCM_Menu_4 & "," & RCM_Menu_5 & "," & RCM_Menu_6 & "," & RCM_Menu_7 & "," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & "," & RCM_Menu_14 & "," & RCM_Menu_15 & ",'" & RCM_Menu_16 & "','" & RCM_Menu_17 & "','" & RCM_Menu_18 & "'," & RCM_Menu_19 & "," & RCM_Menu_20 & "," & RCM_Menu_21 & ",'" & RCM_Menu_22 & "']);"
    Else
        stm_bpx = "stm_bpx('" & bpOID & "','" & bpTID & "',[1,2,-2,-3," & RCM_Menu_4 & "," & RCM_Menu_5 & ",0," & RCM_Menu_7 & "," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & "," & RCM_Menu_14 & "," & RCM_Menu_15 & ",'" & RCM_Menu_16 & "','" & RCM_Menu_17 & "','" & RCM_Menu_18 & "'," & RCM_Menu_19 & "," & RCM_Menu_20 & "," & RCM_Menu_21 & ",'" & RCM_Menu_22 & "']);"
    End If
End Function

Function stm_ai()
    stm_ai = "stm_ai('p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0," & RCM_Item_22 & "," & RCM_Item_23 & ",'" & RCM_Item_24 & "'," & RCM_Item_25 & ",'" & RCM_Item_26 & "'," & RCM_Item_27 & ",'" & RCM_Item_28 & "','" & RCM_Item_29 & "'," & RCM_Item_30 & "," & RCM_Item_31 & "," & RCM_Item_32 & "," & RCM_Item_33 & ",'" & RCM_Item_34 & "','" & RCM_Item_35 & "','" & RCM_Item_36 & "','" & RCM_Item_37 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "']);"
End Function

Function stm_aix(mOID, mTID, mClassName, mClassFile, mOpenType, mMenuTitle, mSubClass)
    If mSubClass = False Then
        stm_aix = "stm_aix('" & mOID & "','" & mTID & "',[0,'" & mClassName & "','','',-1,-1,0,'" & mClassFile & "','" & mOpenType & "','" & mClassFile & "','" & EncodeJS(mMenuTitle) & "','','',0,0,0,'','',0,0,0," & RCM_Item_22 & "," & RCM_Item_23 & ",'" & RCM_Item_24 & "'," & RCM_Item_25 & ",'" & RCM_Item_26 & "'," & RCM_Item_27 & ",'" & RCM_Item_28 & "','" & RCM_Item_29 & "'," & RCM_Item_30 & "," & RCM_Item_31 & "," & RCM_Item_32 & "," & RCM_Item_33 & ",'" & RCM_Item_34 & "','" & RCM_Item_35 & "','" & RCM_Item_36 & "','" & RCM_Item_37 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "']);"
    ElseIf mSubClass = True Then
        stm_aix = "stm_aix('" & mOID & "','" & mTID & "',[0,'" & mClassName & "','','',-1,-1,0,'" & mClassFile & "','" & mOpenType & "','" & mClassFile & "','" & EncodeJS(mMenuTitle) & "','','',6,0,0,'" & strInstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0," & RCM_Item_22 & "," & RCM_Item_23 & ",'" & RCM_Item_24 & "'," & RCM_Item_25 & ",'" & RCM_Item_26 & "'," & RCM_Item_27 & ",'" & RCM_Item_28 & "','" & RCM_Item_29 & "'," & RCM_Item_30 & "," & RCM_Item_31 & "," & RCM_Item_32 & "," & RCM_Item_33 & ",'" & RCM_Item_34 & "','" & RCM_Item_35 & "','" & RCM_Item_36 & "','" & RCM_Item_37 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "']);"
    End If
End Function
    
Function EncodeJS(str)
    EncodeJS = Replace(Replace(Replace(Replace(Replace(str, Chr(10), ""), "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "")
End Function

Sub ShowDemoMenu()
    Response.Write "<script type='text/javascript' language='JavaScript1.2' src='" & strInstallDir & "js/stm31.js'></script>"
    Response.Write "<script language='JavaScript'>"
    Response.Write stm_bm() & vbCrLf
    Response.Write stm_bp_h() & vbCrLf
    Response.Write stm_ai() & vbCrLf
    Response.Write stm_aix("p0i1", "p0i0", "Ƶ��������ҳ", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i3", "p0i0", "ѧϰ����", "#", "_self", "", False) & vbCrLf
    Response.Write stm_bp_v("p1") & vbCrLf
    Response.Write stm_aix("p1i0", "p0i0", "���ݿ�����", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p1i1", "p0i0", "ASP����", "#", "_self", "", True) & vbCrLf
    Response.Write stm_bpx("p2", "p1", 1) & vbCrLf
    Response.Write stm_aix("p2i0", "p1i0", "��̼���", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i1", "p1i0", "����Դ��", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i2", "p1i0", "�����ϼ�", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i3", "p1i0", "�﷨�ٲ�", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p2i2", "p1i0", "�������", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i3", "p1i0", "����������", "#", "_self", "", True) & vbCrLf
    Response.Write stm_bpx("p3", "p2", 1) & vbCrLf
    Response.Write stm_aix("p3i0", "p2i0", "WEB������", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i1", "p2i0", "FTP������", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p3i4", "p2i0", "���簲ȫ", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i5", "p2i0", "��������", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i4", "p0i0", "������", "#", "_self", "", False) & vbCrLf
    Response.Write stm_bpx("p3", "p2", 0) & vbCrLf
    Response.Write stm_aix("p3i0", "p2i0", "PHP���", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i1", "p2i0", "JSP���", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i2", "p2i0", ".NET���", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i5", "p0i0", "�����鼮", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
    Response.Write "stm_em();" & vbCrLf
    Response.Write "</script>"
End Sub

Function FilterString(strChar)
    If strChar = "" Or IsNull(strChar) Then
        FilterString = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "',%,<,>," & Chr(34) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    FilterString = tempChar
End Function
%>
