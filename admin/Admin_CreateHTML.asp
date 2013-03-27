<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 0      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Response.Write "<html><head><title>" & ChannelName & "生成管理</title>" & vbCrLf
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
Call ShowPageTitle(ChannelName & "生成管理", 10008)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>生成说明：</strong></td>" & vbCrLf
Response.Write "    <td>生成操作比较消耗系统资源及费时，每次生成时，请尽量减少要生成的文件量。"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Call PopCalendarInit
If Action = "SiteSpecial" Then
    Response.Write "<br><table width='100%' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>"
    Response.Write "    <td align='center'>全站专题生成管理</td>"
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
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='1'"" value='生成选定专题的列表页' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='2'"" value='生成所有专题的列表页' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            提示：<br>"
    Response.Write "            可以按住“CTRL”或“Shift”键进行多选"
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
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>内容页生成</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>栏目页生成</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>专题页生成</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>其他页生成</td>" & vbCrLf
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
    Response.Write "            生成最新 <input name='TopNew' id='TopNew' value='50' size=8 maxlength='10'> 篇" & ChannelShortName & "&nbsp;"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='4'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form5' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            生成更新时间为"
    Response.Write "            <input name='BeginDate' type='text' id='BeginDate' value='" & FormatDateTime(Date, 2) & "'  size=10 maxlength='20'> "
    Response.Write "            <a style='cursor:hand;' onClick='PopCalendar.show(document.form5.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
    Response.Write "            到"
    Response.Write "            <input name='EndDate' type='text' id='EndDate' value='" & FormatDateTime(Date, 2) & "'  size=10 maxlength='20' title='不包含此日期'>"
    Response.Write "            <a style='cursor:hand;' onClick='PopCalendar.show(document.form5.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
    Response.Write "            的" & ChannelShortName & ""
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='5'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form6' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            生成ID号为 <input name='BeginID' type='text' id='BeginID' value='1' size=8 maxlength='10'> 到 <input name='EndID' type='text' id='EndID' value='100' size=8 maxlength='10'> 的" & ChannelShortName & ""
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='6'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form1' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            生成指定ID的" & ChannelShortName & "（多个ID可用逗号隔开）："
    Response.Write "            <input name='" & ModuleName & "ID' type='text' id='" & ModuleName & "ID' value='1,3,5' size='20'>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='1'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form2' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            生成指定栏目的" & ChannelShortName & "："
    Response.Write "            <select name='ClassID'>" & GetClass_Option(5, 0) & "</select>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='2'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form9' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            <font color='red'>生成所有未生成的" & ChannelShortName & "（推荐）</font>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='9'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='开始生成>>'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_UpdateCreatedStatus.asp?ChannelID=" & ChannelID & "' title='当您用FTP工具删除了已经生成的HTML页面，或者换了服务器后，可以用此功能更新数据库记录的生成状态，以能正常使用“生成所有未生成的" & ChannelShortName & "”功能。'>【更新" & ChannelShortName & "生成状态】</a>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form3' method='post' action='Admin_Create" & ModuleName & ".asp'>"
    Response.Write "            生成所有" & ChannelShortName & ""
    Response.Write "            <input name='Action' type='hidden' id='Action' value='Create" & ModuleName & "'>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='3'>"
    Response.Write "            <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "            <input name='submit' type='submit' id='submit' value='开始生成>>'>"
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
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formclass.CreateType.value='1'"" value='生成选定栏目的列表页' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formclass.CreateType.value='2'"" value='生成所有栏目的列表页' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            提示：<br>"
    Response.Write "            可以按住“CTRL”或“Shift”键进行多选"
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
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='1'"" value='生成选定专题的列表页' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <input type='submit' name='Submit' onClick=""document.formspecial.CreateType.value='2'"" value='生成所有专题的列表页' style='cursor:hand;'>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            <br>"
    Response.Write "            提示：<br>"
    Response.Write "            可以按住“CTRL”或“Shift”键进行多选"
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
    Response.Write "      <input name='submit' type='submit' id='submit' value=' 生成" & ChannelName & "首页 '>"
    Response.Write "      </form>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</td></tr></table>"
End If
Response.Write "</body></html>"
Call CloseConn

%>
