<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 0      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Response.Write "<html><head><title>" & ChannelName & "���ɹ���</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
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
Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle(ChannelName & "���ɹ���", 10008)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>����˵����</strong></td>" & vbCrLf
Response.Write "    <td>���ɲ����Ƚ�����ϵͳ��Դ����ʱ��ÿ������ʱ���뾡������Ҫ���ɵ��ļ�����"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Call PopCalendarInit
If Action = "SiteSpecial" Then
    Response.Write "<br><table width='100%' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>"
    Response.Write "    <td align='center'>ȫվר�����ɹ���</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td align='center' class='tdbg'>"
    Response.Write "      <table width='530' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "        <form name='formspecial' method='post' action='Admin_CreateSiteSpecial.asp'>"
    Response.Write "        <tr>"
    Response.Write "          <td>"
    Response.Write "          <select name='SpecialID' size='2' multiple id='SpecialID' style='height:300px;width:300px;'>" & GetSpecial_Option(0) & "</select>"
    Response.Write "          </td>"
    Response.Write "          <td valign='bottom'>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='CreateSiteSpecial'>"
    Response.Write "            <input name='CreateType' type='hidden' value='1'>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='1'"" value='����ѡ��ר����б�ҳ' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='2'"" value='��������ר����б�ҳ' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            ��ʾ��<br>"
    Response.Write "            ���԰�ס��CTRL����Shift�������ж�ѡ"
    Response.Write "            <br>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        </form>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf
Else
    Response.Write "<br>"
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>����ҳ����</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��Ŀҳ����</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>ר��ҳ����</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>����ҳ����</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "  <tr>"
    Response.Write "    <td class='tdbg'>"
    Response.Write "      <table width='500' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form4' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            �������� <input name='TopNew' id='TopNew' value='50' size=8 maxlength='10'> ƪ" & ChannelShortName & "&nbsp;"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='4'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='��ʼ����>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form5' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            ���ɸ���ʱ��Ϊ"
    Response.Write "            <input name='BeginDate' type='text' id='BeginDate' value='" & FormatDateTime(Date, 2) & "'  size=10 maxlength='20'> "
    Response.Write "            <a style='cursor:hand;' onClick='PopCalendar.show(document.form5.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
    Response.Write "            ��"
    Response.Write "            <input name='EndDate' type='text' id='EndDate' value='" & FormatDateTime(Date, 2) & "'  size=10 maxlength='20' title='������������'>"
    Response.Write "            <a style='cursor:hand;' onClick='PopCalendar.show(document.form5.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
    Response.Write "            ��" & ChannelShortName & ""
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='5'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='��ʼ����>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form6' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            ����ID��Ϊ <input name='BeginID' type='text' id='BeginID' value='1' size=8 maxlength='10'> �� <input name='EndID' type='text' id='EndID' value='100' size=8 maxlength='10'> ��" & ChannelShortName & ""
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='6'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='��ʼ����>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form1' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            ����ָ��ID��" & ChannelShortName & "�����ID���ö��Ÿ�������"
    Response.Write "            <input name='" & ModuleName & "ID' type='text' id='" & ModuleName & "ID' value='1,3,5' size='20'>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='1'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='��ʼ����>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form2' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            ����ָ����Ŀ��" & ChannelShortName & "��"
    Response.Write "            <select name='ClassID'>" & GetClass_Option(5, 0) & "</select>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='2'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='��ʼ����>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form9' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            <font color='red'>��������δ���ɵ�" & ChannelShortName & "���Ƽ���</font>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='9'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='��ʼ����>>'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_UpdateCreatedStatus.asp?ChannelID=" & ChannelID & "' title='������FTP����ɾ�����Ѿ����ɵ�HTMLҳ�棬���߻��˷������󣬿����ô˹��ܸ������ݿ��¼������״̬����������ʹ�á���������δ���ɵ�" & ChannelShortName & "�����ܡ�'>������" & ChannelShortName & "����״̬��</a>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form3' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            ��������" & ChannelShortName & ""
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='3'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='��ʼ����>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </tbody>"


    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "  <tr>"
    Response.Write "    <td align='center' class='tdbg'>"
    Response.Write "      <table width='530' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "        <form name='formclass' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "        <tr>"
    Response.Write "          <td>"
    Response.Write "            <select name='ClassID' size='2' multiple style='height:300px;width:300px;'>" & GetClass_Option(5, 0) & "</select>"
    Response.Write "          </td>"
    Response.Write "          <td valign='bottom'>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='CreateClass'>"
    Response.Write "            <input name='CreateType' type='hidden' value='1'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formclass.CreateType.value='1'"" value='����ѡ����Ŀ���б�ҳ' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formclass.CreateType.value='2'"" value='����������Ŀ���б�ҳ' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            ��ʾ��<br>"
    Response.Write "            ���԰�ס��CTRL����Shift�������ж�ѡ"
    Response.Write "            <br>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        </form>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </tbody>"


    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "  <tr>"
    Response.Write "    <td align='center' class='tdbg'>"
    Response.Write "      <table width='530' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "        <form name='formspecial' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "        <tr>"
    Response.Write "          <td>"
    Response.Write "          <select name='SpecialID' size='2' multiple id='SpecialID' style='height:300px;width:300px;'>" & GetSpecial_Option(0) & "</select>"
    Response.Write "          </td>"
    Response.Write "          <td valign='bottom'>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='CreateSpecial'>"
    Response.Write "            <input name='CreateType' type='hidden' value='1'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='1'"" value='����ѡ��ר����б�ҳ' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='2'"" value='��������ר����б�ҳ' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            ��ʾ��<br>"
    Response.Write "            ���԰�ס��CTRL����Shift�������ж�ѡ"
    Response.Write "            <br>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        </form>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </tbody>"


    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "  <tr>"
    Response.Write "    <td align='center' class='tdbg'>"
    Response.Write "      <form name='form' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "      <br>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='CreateIndex'>"
    Response.Write "      <input name='CreateType' type='hidden' id='CreateType' value='1'>"
    Response.Write "      <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "      <input name='submit' type='submit' id='submit' value=' ����" & ChannelName & "��ҳ '>"
    Response.Write "      </form>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</td></tr></table>"
End If
Response.Write "</body></html>"
Call CloseConn

%>
