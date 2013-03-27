<!--#include file="Admin_Common.asp"-->
<!--#include file="RootClass_Menu_Config.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 1   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

If ChannelID = 0 Then
    Response.Write "频道参数错！"
    Response.End
End If

If AdminPurview > 1 And CheckPurview_Other(AdminPurview_Others, "Menu_" & ChannelDir) = False Then
    Response.Write "你没有此项操作的权限！"
    Response.End
End If

Dim strTopMenu, pNum, pNum2, OpenType_Class, strMenuJS

Response.Write "<html><head><title>顶部栏目菜单管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Response.Write "  <tr class='topbg'> " & vbCrLf
Response.Write "    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>顶部栏目菜单生成</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=10013' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>"
Response.Write "    <td height='30' colspan='2'>"
Response.Write "<a href='Admin_RootClass_Menu.asp?Action=ShowConfig&ChannelID=" & ChannelID & "' target=main>顶部栏目菜单参数设置</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_RootClass_Menu.asp?Action=ShowCreate&ChannelID=" & ChannelID & "' target=main>顶部栏目菜单生成</a>"
Response.Write "    </td>"
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>"
Response.Write "    <td width='70' height='30'><strong>菜单演示：</strong></td>"
Response.Write "    <td height='30'>"
Call ShowDemoMenu
Response.Write "    </td>"
Response.Write "    <td width='350'>注：参数设置，▲代表鼠标悬停时效果，▼代表鼠标移出时效果。</td>"
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
    Response.Write "    <td height='22' colspan='6'><strong>顶部栏目菜单参数设置</strong> （注：部分特效只对特定的浏览器才有效）</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>弹出方式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_1' id='RCM_Menu_1'>"
    Response.Write "        <option value='1' "
    If RCM_Menu_1 = "1" Then Response.Write " selected"
    Response.Write "        >向左</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_1 = "2" Then Response.Write " selected"
    Response.Write "        >向右</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_1 = "3" Then Response.Write " selected"
    Response.Write "        >向上</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_1 = "4" Then Response.Write " selected"
    Response.Write "        >向下</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>横向偏移量：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_2' type='text' id='RCM_Menu_2' value='" & RCM_Menu_2 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>纵向偏移量：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_3' type='text' id='RCM_Menu_3' value='" & RCM_Menu_3 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单项边距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_4' type='text' id='RCM_Menu_4' value='" & RCM_Menu_4 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项间距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_5' type='text' id='RCM_Menu_5' value='" & RCM_Menu_5 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项左边距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_6' type='text' id='RCM_Menu_6' value='" & RCM_Menu_6 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单项右边距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_7' type='text' id='RCM_Menu_7' value='" & RCM_Menu_7 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单透明度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_8' type='text' id='RCM_Menu_8' value='" & RCM_Menu_8 & "' size='10' maxlength='10' title='0-100 完全透明-完全不透明'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单其它特效：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_9' type='text' id='RCM_Menu_9' value='" & RCM_Menu_9 & "' size='10' maxlength='200'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单弹出效果▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_10' id='RCM_Menu_10'>"
    Response.Write "        <option value='-1' "
    If RCM_Menu_10 = "-1" Then Response.Write " selected"
    Response.Write "        >无特效</option>"
    Response.Write "        <option value='0' "
    If RCM_Menu_10 = "0" Then Response.Write " selected"
    Response.Write "        >方形收缩</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_10 = "1" Then Response.Write " selected"
    Response.Write "        >方形扩散</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_10 = "2" Then Response.Write " selected"
    Response.Write "        >圆形收缩</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_10 = "3" Then Response.Write " selected"
    Response.Write "        >圆形扩散</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_10 = "4" Then Response.Write " selected"
    Response.Write "        >上拉效果</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_10 = "5" Then Response.Write " selected"
    Response.Write "        >下拉效果</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_10 = "6" Then Response.Write " selected"
    Response.Write "        >从左向右</option>"
    Response.Write "        <option value='7' "
    If RCM_Menu_10 = "7" Then Response.Write " selected"
    Response.Write "        >从右向左</option>"
    Response.Write "        <option value='8' "
    If RCM_Menu_10 = "8" Then Response.Write " selected"
    Response.Write "        >左右百叶</option>"
    Response.Write "        <option value='9' "
    If RCM_Menu_10 = "9" Then Response.Write " selected"
    Response.Write "        >上下百叶</option>"
    Response.Write "        <option value='10' "
    If RCM_Menu_10 = "10" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='11' "
    If RCM_Menu_10 = "11" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='12' "
    If RCM_Menu_10 = "12" Then Response.Write " selected"
    Response.Write "        >模糊效果</option>"
    Response.Write "        <option value='13' "
    If RCM_Menu_10 = "13" Then Response.Write " selected"
    Response.Write "        >左右关门</option>"
    Response.Write "        <option value='14' "
    If RCM_Menu_10 = "14" Then Response.Write " selected"
    Response.Write "        >左右开门</option>"
    Response.Write "        <option value='15' "
    If RCM_Menu_10 = "15" Then Response.Write " selected"
    Response.Write "        >上下关门</option>"
    Response.Write "        <option value='16' "
    If RCM_Menu_10 = "16" Then Response.Write " selected"
    Response.Write "        >上下开门</option>"
    Response.Write "        <option value='17' "
    If RCM_Menu_10 = "17" Then Response.Write " selected"
    Response.Write "        >左下拉开</option>"
    Response.Write "        <option value='18' "
    If RCM_Menu_10 = "18" Then Response.Write " selected"
    Response.Write "        >左上拉开</option>"
    Response.Write "        <option value='19' "
    If RCM_Menu_10 = "19" Then Response.Write " selected"
    Response.Write "        >右下拉开</option>"
    Response.Write "        <option value='20' "
    If RCM_Menu_10 = "20" Then Response.Write " selected"
    Response.Write "        >右上拉开</option>"
    Response.Write "        <option value='21' "
    If RCM_Menu_10 = "21" Then Response.Write " selected"
    Response.Write "        >上下条纹</option>"
    Response.Write "        <option value='22' "
    If RCM_Menu_10 = "22" Then Response.Write " selected"
    Response.Write "        >左右条纹</option>"
    Response.Write "        <option value='23' "
    If RCM_Menu_10 = "23" Then Response.Write " selected"
    Response.Write "        >随机特效</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单弹出效果▼：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_12' id='RCM_Menu_12'>"
    Response.Write "        <option value='-1' "
    If RCM_Menu_12 = "-1" Then Response.Write " selected"
    Response.Write "        >无特效</option>"
    Response.Write "        <option value='0' "
    If RCM_Menu_12 = "0" Then Response.Write " selected"
    Response.Write "        >方形收缩</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_12 = "1" Then Response.Write " selected"
    Response.Write "        >方形扩散</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_12 = "2" Then Response.Write " selected"
    Response.Write "        >圆形收缩</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_12 = "3" Then Response.Write " selected"
    Response.Write "        >圆形扩散</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_12 = "4" Then Response.Write " selected"
    Response.Write "        >上拉效果</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_12 = "5" Then Response.Write " selected"
    Response.Write "        >下拉效果</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_12 = "6" Then Response.Write " selected"
    Response.Write "        >从左向右</option>"
    Response.Write "        <option value='7' "
    If RCM_Menu_12 = "7" Then Response.Write " selected"
    Response.Write "        >从右向左</option>"
    Response.Write "        <option value='8' "
    If RCM_Menu_12 = "8" Then Response.Write " selected"
    Response.Write "        >左右百叶</option>"
    Response.Write "        <option value='9' "
    If RCM_Menu_12 = "9" Then Response.Write " selected"
    Response.Write "        >上下百叶</option>"
    Response.Write "        <option value='10' "
    If RCM_Menu_12 = "10" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='11' "
    If RCM_Menu_12 = "11" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='12' "
    If RCM_Menu_12 = "12" Then Response.Write " selected"
    Response.Write "        >模糊效果</option>"
    Response.Write "        <option value='13' "
    If RCM_Menu_12 = "13" Then Response.Write " selected"
    Response.Write "        >左右关门</option>"
    Response.Write "        <option value='14' "
    If RCM_Menu_12 = "14" Then Response.Write " selected"
    Response.Write "        >左右开门</option>"
    Response.Write "        <option value='15' "
    If RCM_Menu_12 = "15" Then Response.Write " selected"
    Response.Write "        >上下关门</option>"
    Response.Write "        <option value='16' "
    If RCM_Menu_12 = "16" Then Response.Write " selected"
    Response.Write "        >上下开门</option>"
    Response.Write "        <option value='17' "
    If RCM_Menu_12 = "17" Then Response.Write " selected"
    Response.Write "        >左下拉开</option>"
    Response.Write "        <option value='18' "
    If RCM_Menu_12 = "18" Then Response.Write " selected"
    Response.Write "        >左上拉开</option>"
    Response.Write "        <option value='19' "
    If RCM_Menu_12 = "19" Then Response.Write " selected"
    Response.Write "        >右下拉开</option>"
    Response.Write "        <option value='20' "
    If RCM_Menu_12 = "20" Then Response.Write " selected"
    Response.Write "        >右上拉开</option>"
    Response.Write "        <option value='21' "
    If RCM_Menu_12 = "21" Then Response.Write " selected"
    Response.Write "        >上下条纹</option>"
    Response.Write "        <option value='22' "
    If RCM_Menu_12 = "22" Then Response.Write " selected"
    Response.Write "        >左右条纹</option>"
    Response.Write "        <option value='23' "
    If RCM_Menu_12 = "23" Then Response.Write " selected"
    Response.Write "        >随机特效</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单弹出效果速度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_13' type='text' id='RCM_Menu_13' value='" & RCM_Menu_13 & "' size='10' maxlength='10' title='速度值：10-100'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单阴影效果：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_14' id='RCM_Menu_14'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_14 = "0" Then Response.Write " selected"
    Response.Write "        >无阴影</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_14 = "1" Then Response.Write " selected"
    Response.Write "        >简单阴影</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_14 = "2" Then Response.Write " selected"
    Response.Write "        >复杂阴影</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单阴影深度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_15' type='text' id='RCM_Menu_15' value='" & RCM_Menu_15 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单阴影颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_16' type='text' id='RCM_Menu_16' value='" & RCM_Menu_16 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单背景颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_17' type='text' id='RCM_Menu_17' value='" & RCM_Menu_17 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单背景图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_18' type='text' id='RCM_Menu_18' value='" & RCM_Menu_18 & "' size='10' maxlength='200' title='只有当菜单项背景颜色设为透明色：transparent 时才有效'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>背景图片平铺模式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_19' id='RCM_Menu_19'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_19 = "0" Then Response.Write " selected"
    Response.Write "        >不平铺</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_19 = "1" Then Response.Write " selected"
    Response.Write "        >横向平铺</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_19 = "2" Then Response.Write " selected"
    Response.Write "        >纵向平铺</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_19 = "3" Then Response.Write " selected"
    Response.Write "        >完全平铺</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单边框类型：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_20' id='RCM_Menu_20'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_20 = "0" Then Response.Write " selected"
    Response.Write "        >无边框</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_20 = "1" Then Response.Write " selected"
    Response.Write "        >单实线</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_20 = "2" Then Response.Write " selected"
    Response.Write "        >双实线</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_20 = "5" Then Response.Write " selected"
    Response.Write "        >凹陷</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_20 = "6" Then Response.Write " selected"
    Response.Write "        >凸起</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单边框宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_21' type='text' id='RCM_Menu_21' value='" & RCM_Menu_21 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单边框颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_22' type='text' id='RCM_Menu_22' value='" & RCM_Menu_22 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"

    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='6'><strong>菜单项参数设置</strong></td>"
    Response.Write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>菜单项类型：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <select name='RCM_Item_1' id='RCM_Item_1'>"
'    response.write "        <option value='0' "
'   if RCM_Menu_1="0" then response.write " selected"
'    response.write "        >文本</option>"
'    response.write "        <option value='1' "
'   if RCM_Menu_1="1" then response.write " selected"
'    response.write "        >HTML</option>"
'    response.write "        <option value='2' "
'   if RCM_Menu_1="2" then response.write " selected"
'    response.write "        >图片</option>"
'    response.write "      </select>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>菜单项名称：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_2' type='text' id='RCM_Item_2' value='" & RCM_Item_2 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>图片文件：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_3' type='text' id='RCM_Item_3' value='" & RCM_Item_3 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>鼠标指在菜单项时，图片文件：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_4' type='text' id='RCM_Item_4' value='" & RCM_Item_4 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>图片宽度：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_5' type='text' id='RCM_Item_5' value='" & RCM_Item_5 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>图片高度：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_6' type='text' id='RCM_Item_6' value='" & RCM_Item_6 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>图片边框：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_7' type='text' id='RCM_Item_7' value='" & RCM_Item_7 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>链接地址：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_8' type='text' id='RCM_Item_8' value='" & RCM_Item_8 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>链接目标：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_9' type='text' id='RCM_Item_9' value='" & RCM_Item_9 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>链接状态栏显示：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_10' type='text' id='RCM_Item_10' value='" & RCM_Item_10 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>链接地址提示信息：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_11' type='text' id='RCM_Item_11' value='" & RCM_Item_11 & "' size='10' maxlength='10'>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong></strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      "
'    response.write "    </td>"
'    response.write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单项左图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_12' type='text' id='RCM_Item_12' value='" & RCM_Item_12 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项左图片▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_13' type='text' id='RCM_Item_13' value='" & RCM_Item_13 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>左图片宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_14' type='text' id='RCM_Item_14' value='" & RCM_Item_14 & "' size='10' maxlength='10' title='0为图像原始宽度'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>左图片高度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_15' type='text' id='RCM_Item_15' value='" & RCM_Item_15 & "' size='10' maxlength='10' title='0为图像原始高度'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>左图片边框大小：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_16' type='text' id='RCM_Item_16' value='" & RCM_Item_16 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项右图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_17' type='text' id='RCM_Item_17' value='" & RCM_Item_17 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单项右图片▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_18' type='text' id='RCM_Item_18' value='" & RCM_Item_18 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>右图片宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_19' type='text' id='RCM_Item_19' value='" & RCM_Item_19 & "' size='10' maxlength='10' title='0为图像原始宽度'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>右图片高度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_20' type='text' id='RCM_Item_20' value='" & RCM_Item_20 & "' size='10' maxlength='10' title='0为图像原始高度'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>右图片边框大小：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_21' type='text' id='RCM_Item_21' value='" & RCM_Item_21 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>文字水平对齐方式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_22' id='RCM_Item_22'>"
    Response.Write "        <option value='0' "
    If RCM_Item_22 = "0" Then Response.Write " selected"
    Response.Write "        >左对齐</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_22 = "1" Then Response.Write " selected"
    Response.Write "        >居中</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_22 = "2" Then Response.Write " selected"
    Response.Write "        >右对齐</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>文字垂直对齐方式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_23' id='RCM_Item_23'>"
    Response.Write "        <option value='0' "
    If RCM_Item_23 = "0" Then Response.Write " selected"
    Response.Write "        >顶部</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_23 = "1" Then Response.Write " selected"
    Response.Write "        >居中</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_23 = "2" Then Response.Write " selected"
    Response.Write "        >底部</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单项背景颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_24' type='text' id='RCM_Item_24' value='" & RCM_Item_24 & "' size='10' maxlength='10' title='透明色：transparent'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>背景颜色是否显示：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_25' id='RCM_Item_25'>"
    Response.Write "        <option value='0' "
    If RCM_Item_25 = "0" Then Response.Write " selected"
    Response.Write "        >显示</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_25 = "1" Then Response.Write " selected"
    Response.Write "        >不显示</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项背景颜色▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_26' type='text' id='RCM_Item_26' value='" & RCM_Item_26 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>背景颜色是否显示▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_27' id='RCM_Item_27'>"
    Response.Write "        <option value='0' "
    If RCM_Item_27 = "0" Then Response.Write " selected"
    Response.Write "        >显示</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_27 = "1" Then Response.Write " selected"
    Response.Write "        >不显示</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项背景图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_28' type='text' id='RCM_Item_28' value='" & RCM_Item_28 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项背景图片▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_29' type='text' id='RCM_Item_29' value='" & RCM_Item_29 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>背景图片平铺模式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_30' id='RCM_Item_30'>"
    Response.Write "        <option value='0' "
    If RCM_Item_30 = "0" Then Response.Write " selected"
    Response.Write "        >不平铺</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_30 = "1" Then Response.Write " selected"
    Response.Write "        >横向平铺</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_30 = "2" Then Response.Write " selected"
    Response.Write "        >纵向平铺</option>"
    Response.Write "        <option value='3' "
    If RCM_Item_30 = "3" Then Response.Write " selected"
    Response.Write "        >完全平铺</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
'    response.write "    <td width='130' height='25'><strong>背景图片平铺模式▲：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <select name='RCM_Item_31' id='RCM_Item_31'>"
'    response.write "        <option value='0' "
'   if RCM_Menu_1="0" then response.write " selected"
'    response.write "        >不平铺</option>"
'    response.write "        <option value='1' "
'   if RCM_Menu_1="1" then response.write " selected"
'    response.write "        >横向平铺</option>"
'    response.write "        <option value='2' "
'   if RCM_Menu_1="2" then response.write " selected"
'    response.write "        >纵向平铺</option>"
'    response.write "        <option value='3' "
'   if RCM_Menu_1="3" then response.write " selected"
'    response.write "        >完全平铺</option>"
'    response.write "      </select>"
'    response.write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项边框类型：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_32' id='RCM_Item_32'>"
    Response.Write "        <option value='0' "
    If RCM_Item_32 = "0" Then Response.Write " selected"
    Response.Write "        >无边框</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_32 = "1" Then Response.Write " selected"
    Response.Write "        >单实线</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_32 = "2" Then Response.Write " selected"
    Response.Write "        >双实线</option>"
    Response.Write "        <option value='5' "
    If RCM_Item_32 = "5" Then Response.Write " selected"
    Response.Write "        >凹陷</option>"
    Response.Write "        <option value='6' "
    If RCM_Item_32 = "6" Then Response.Write " selected"
    Response.Write "        >凸起</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项边框宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_33' type='text' id='RCM_Item_33' value='" & RCM_Item_33 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单项边框颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_34' type='text' id='RCM_Item_34' value='" & RCM_Item_34 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项边框颜色▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_35' type='text' id='RCM_Item_35' value='" & RCM_Item_35 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项文字颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_36' type='text' id='RCM_Item_36' value='" & RCM_Item_36 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25'><strong>菜单项文字颜色▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_37' type='text' id='RCM_Item_37' value='" & RCM_Item_37 & "' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25'><strong>菜单项文字字体：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='FontName_RCM_Item_38' id='FontName_RCM_Item_38'>"
    Response.Write "        <option value='宋体' "
    If FontName_RCM_Item_38 = "宋体" Then Response.Write " selected"
    Response.Write "        >宋体</option>"
    Response.Write "        <option value=""黑体"" "
    If FontName_RCM_Item_38 = "黑体" Then Response.Write " selected"
    Response.Write "        >黑体</option>"
    Response.Write "        <option value=""楷体_GB2312"" "
    If FontName_RCM_Item_38 = "楷体_GB2312" Then Response.Write " selected"
    Response.Write "        >楷体</option>"
    Response.Write "        <option value=""仿宋_GB2312"" "
    If FontName_RCM_Item_38 = "仿宋_GB2312" Then Response.Write " selected"
    Response.Write "        >仿宋</option>"
    Response.Write "        <option value=""隶书"" "
    If FontName_RCM_Item_38 = "隶书" Then Response.Write " selected"
    Response.Write "        >隶书</option>"
    Response.Write "        <option value=""幼圆"" "
    If FontName_RCM_Item_38 = "幼圆" Then Response.Write " selected"
    Response.Write "        >幼圆</option>"
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
    Response.Write "    <td width='130' height='25'><strong>菜单项文字字体▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='FontName_RCM_Item_39' id='FontName_RCM_Item_39'>"
    Response.Write "        <option value='宋体' "
    If FontName_RCM_Item_39 = "宋体" Then Response.Write " selected"
    Response.Write "        >宋体</option>"
    Response.Write "        <option value=""黑体"" "
    If FontName_RCM_Item_39 = "黑体" Then Response.Write " selected"
    Response.Write "        >黑体</option>"
    Response.Write "        <option value=""楷体_GB2312"" "
    If FontName_RCM_Item_39 = "楷体_GB2312" Then Response.Write " selected"
    Response.Write "        >楷体</option>"
    Response.Write "        <option value=""仿宋_GB2312"" "
    If FontName_RCM_Item_39 = "仿宋_GB2312" Then Response.Write " selected"
    Response.Write "        >仿宋</option>"
    Response.Write "        <option value=""隶书"" "
    If FontName_RCM_Item_39 = "隶书" Then Response.Write " selected"
    Response.Write "        >隶书</option>"
    Response.Write "        <option value=""幼圆"" "
    If FontName_RCM_Item_39 = "幼圆" Then Response.Write " selected"
    Response.Write "        >幼圆</option>"
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
    Response.Write "      <input name='cmdSave' type='submit' id='cmdSave' value=' 保存设置 ' "
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
        ErrMsg = ErrMsg & "<br><li>你的服务器不支持 FSO(Scripting.FileSystemObject)! </li>"
        Exit Sub
    End If
    Set hf = fso.CreateTextFile(Server.MapPath(InstallDir & AdminDir & "/RootClass_Menu_Config.asp"), True)

    hf.Write "<" & "%" & vbCrLf
    hf.Write "'菜单显示参数设置" & vbCrLf
    hf.Write "Const RCM_Menu_1=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_1"))) & Chr(34) & "      '菜单弹出方式 1：左  2：右  3：上  4：下" & vbCrLf
    hf.Write "Const RCM_Menu_2=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_2"))) & Chr(34) & "      '菜单弹出横向偏移量" & vbCrLf
    hf.Write "Const RCM_Menu_3=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_3"))) & Chr(34) & "      '菜单弹出纵向偏移量" & vbCrLf
    hf.Write "Const RCM_Menu_4=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_4"))) & Chr(34) & "      '菜单项边距" & vbCrLf
    hf.Write "Const RCM_Menu_5=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_5"))) & Chr(34) & "      '菜单项间距" & vbCrLf
    hf.Write "Const RCM_Menu_6=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_6"))) & Chr(34) & "      '菜单项左边距" & vbCrLf
    hf.Write "Const RCM_Menu_7=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_7"))) & Chr(34) & "      '菜单项右边距" & vbCrLf
    hf.Write "Const RCM_Menu_8=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_8"))) & Chr(34) & "      '菜单透明度         0-100 完全透明-完全不透明" & vbCrLf
    hf.Write "Const RCM_Menu_9=" & Chr(34) & FilterString(Trim(request("RCM_Menu_9"))) & Chr(34) & "      '其它特效" & vbCrLf
    hf.Write "Const RCM_Menu_10=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_10"))) & Chr(34) & "        '鼠标指在菜单项时，菜单弹出效果" & vbCrLf
    hf.Write "Const RCM_Menu_11=" & Chr(34) & FilterString(Trim(request("RCM_Menu_11"))) & Chr(34) & "        '其它特效" & vbCrLf
    hf.Write "Const RCM_Menu_12=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_12"))) & Chr(34) & "        '鼠标移出菜单项时，菜单弹出效果" & vbCrLf
    hf.Write "Const RCM_Menu_13=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_13"))) & Chr(34) & "        '菜单弹出效果速度  10-100" & vbCrLf
    hf.Write "Const RCM_Menu_14=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_14"))) & Chr(34) & "        '弹出菜单阴影效果 0：none  1：simple  2：complex" & vbCrLf
    hf.Write "Const RCM_Menu_15=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_15"))) & Chr(34) & "        '弹出菜单阴影深度" & vbCrLf
    hf.Write "Const RCM_Menu_16=" & Chr(34) & FilterString(Trim(request("RCM_Menu_16"))) & Chr(34) & "        '弹出菜单阴影颜色" & vbCrLf
    hf.Write "Const RCM_Menu_17=" & Chr(34) & FilterString(Trim(request("RCM_Menu_17"))) & Chr(34) & "        '弹出菜单背景颜色" & vbCrLf
    hf.Write "Const RCM_Menu_18=" & Chr(34) & FilterString(Trim(request("RCM_Menu_18"))) & Chr(34) & "        '弹出菜单背景图片，只有当菜单项背景颜色设为透明色：transparent 时才有效" & vbCrLf
    hf.Write "Const RCM_Menu_19=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_19"))) & Chr(34) & "        '弹出菜单背景图片平铺模式。 0：不平铺  1：横向平铺  2：纵向平铺  3：完全平铺" & vbCrLf
    hf.Write "Const RCM_Menu_20=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_20"))) & Chr(34) & "        '弹出菜单边框类型 0：无边框  1：单实线  2：双实线  5：凹陷  6：凸起" & vbCrLf
    hf.Write "Const RCM_Menu_21=" & Chr(34) & PE_CLng(Trim(request("RCM_Menu_21"))) & Chr(34) & "        '弹出菜单边框宽度" & vbCrLf
    hf.Write "Const RCM_Menu_22=" & Chr(34) & FilterString(Trim(request("RCM_Menu_22"))) & Chr(34) & "        '弹出菜单边框颜色" & vbCrLf
    hf.Write "Const RCM_Menu_23=" & Chr(34) & "#ffffff" & Chr(34) & "" & vbCrLf
    hf.Write "" & vbCrLf
    hf.Write "'菜单项参数设置" & vbCrLf
    hf.Write "Const RCM_Item_1=" & Chr(34) & "0" & Chr(34) & "      '菜单项类型  0--Txt  1--Html  2--Image" & vbCrLf
    hf.Write "Const RCM_Item_2=" & Chr(34) & "" & Chr(34) & "       '菜单项名称" & vbCrLf
    hf.Write "Const RCM_Item_3=" & Chr(34) & "" & Chr(34) & "       '菜单项为Image，图片文件" & vbCrLf
    hf.Write "Const RCM_Item_4=" & Chr(34) & "" & Chr(34) & "       '菜单项为Image，鼠标指在菜单项时，图片文件。" & vbCrLf
    hf.Write "Const RCM_Item_5=" & Chr(34) & "-1" & Chr(34) & "     '菜单项为Image，图片宽度" & vbCrLf
    hf.Write "Const RCM_Item_6=" & Chr(34) & "-1" & Chr(34) & "     '菜单项为Image，图片高度" & vbCrLf
    hf.Write "Const RCM_Item_7=" & Chr(34) & "0" & Chr(34) & "      '菜单项为Image，图片边框" & vbCrLf
    hf.Write "Const RCM_Item_8=" & Chr(34) & "" & Chr(34) & "       '菜单项链接地址" & vbCrLf
    hf.Write "Const RCM_Item_9=" & Chr(34) & "" & Chr(34) & "       '菜单项链接目标 如：_self  _blank" & vbCrLf
    hf.Write "Const RCM_Item_10=" & Chr(34) & "" & Chr(34) & "      '菜单项链接状态栏显示" & vbCrLf
    hf.Write "Const RCM_Item_11=" & Chr(34) & "" & Chr(34) & "      '菜单项链接地址提示信息" & vbCrLf
    hf.Write "Const RCM_Item_12=" & Chr(34) & FilterString(Trim(request("RCM_Item_12"))) & Chr(34) & "        '菜单项左图片" & vbCrLf
    hf.Write "Const RCM_Item_13=" & Chr(34) & FilterString(Trim(request("RCM_Item_13"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项左图片" & vbCrLf
    hf.Write "Const RCM_Item_14=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_14"))) & Chr(34) & "        '菜单项左图片宽度，0为图像文件原始值" & vbCrLf
    hf.Write "Const RCM_Item_15=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_15"))) & Chr(34) & "        '菜单项左图片高度，0为图像文件原始值" & vbCrLf
    hf.Write "Const RCM_Item_16=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_16"))) & Chr(34) & "        '菜单项左图片边框大小" & vbCrLf
    hf.Write "Const RCM_Item_17=" & Chr(34) & FilterString(Trim(request("RCM_Item_17"))) & Chr(34) & "        '菜单项右图片。如：arrow_r.gif" & vbCrLf
    hf.Write "Const RCM_Item_18=" & Chr(34) & FilterString(Trim(request("RCM_Item_18"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项右图片。如：arrow_w.gif" & vbCrLf
    hf.Write "Const RCM_Item_19=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_19"))) & Chr(34) & "        '菜单项右图片宽度，0为图像文件原始值" & vbCrLf
    hf.Write "Const RCM_Item_20=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_20"))) & Chr(34) & "        '菜单项右图片高度，0为图像文件原始值" & vbCrLf
    hf.Write "Const RCM_Item_21=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_21"))) & Chr(34) & "        '菜单项右图片边框大小" & vbCrLf
    hf.Write "Const RCM_Item_22=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_22"))) & Chr(34) & "        '菜单项文字水平对齐方式  0：左对齐  1：居中  2：右对齐" & vbCrLf
    hf.Write "Const RCM_Item_23=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_23"))) & Chr(34) & "        '菜单项文字垂直对齐方式  0：顶部  1：居中  2：底部" & vbCrLf
    hf.Write "Const RCM_Item_24=" & Chr(34) & FilterString(Trim(request("RCM_Item_24"))) & Chr(34) & "        '菜单项背景颜色  透明色：'transparent'" & vbCrLf
    hf.Write "Const RCM_Item_25=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_25"))) & Chr(34) & "        '菜单项背景颜色是否显示  0：显示  其它：不显示" & vbCrLf
    hf.Write "Const RCM_Item_26=" & Chr(34) & FilterString(Trim(request("RCM_Item_26"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项背景颜色" & vbCrLf
    hf.Write "Const RCM_Item_27=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_27"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项背景颜色是否显示。  0：显示  其它：不显示" & vbCrLf
    hf.Write "Const RCM_Item_28=" & Chr(34) & FilterString(Trim(request("RCM_Item_28"))) & Chr(34) & "        '菜单项背景图片" & vbCrLf
    hf.Write "Const RCM_Item_29=" & Chr(34) & FilterString(Trim(request("RCM_Item_29"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项背景图片" & vbCrLf
    hf.Write "Const RCM_Item_30=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_30"))) & Chr(34) & "        '菜单项背景图片平铺模式。 0：不平铺  1：横向平铺  2：纵向平铺  3：完全平铺" & vbCrLf
    hf.Write "Const RCM_Item_31=" & Chr(34) & "3" & Chr(34) & "     '鼠标指在菜单项时，菜单项背景图片平铺模式。0-3" & vbCrLf
    hf.Write "Const RCM_Item_32=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_32"))) & Chr(34) & "        '菜单项边框类型 0：无边框  1：单实线  2：双实线  5：凹陷  6：凸起" & vbCrLf
    hf.Write "Const RCM_Item_33=" & Chr(34) & PE_CLng(Trim(request("RCM_Item_33"))) & Chr(34) & "        '菜单项边框宽度" & vbCrLf
    hf.Write "Const RCM_Item_34=" & Chr(34) & FilterString(Trim(request("RCM_Item_34"))) & Chr(34) & "        '菜单项边框颜色" & vbCrLf
    hf.Write "Const RCM_Item_35=" & Chr(34) & FilterString(Trim(request("RCM_Item_35"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项边框颜色" & vbCrLf
    hf.Write "Const RCM_Item_36=" & Chr(34) & FilterString(Trim(request("RCM_Item_36"))) & Chr(34) & "        '菜单项文字颜色" & vbCrLf
    hf.Write "Const RCM_Item_37=" & Chr(34) & FilterString(Trim(request("RCM_Item_37"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项文字颜色" & vbCrLf
    hf.Write "Const FontSize_RCM_Item_38=" & Chr(34) & FilterString(Trim(request("FontSize_RCM_Item_38"))) & Chr(34) & "        '菜单项文字大小" & vbCrLf
    hf.Write "Const FontName_RCM_Item_38=" & Chr(34) & FilterString(Trim(request("FontName_RCM_Item_38"))) & Chr(34) & "        '菜单项文字字体" & vbCrLf
    hf.Write "Const FontSize_RCM_Item_39=" & Chr(34) & FilterString(Trim(request("FontSize_RCM_Item_39"))) & Chr(34) & "        '鼠标指在菜单项时,菜单项文字大小" & vbCrLf
    hf.Write "Const FontName_RCM_Item_39=" & Chr(34) & FilterString(Trim(request("FontName_RCM_Item_39"))) & Chr(34) & "        '鼠标指在菜单项时,菜单项文字字体" & vbCrLf
    hf.Write "%" & ">"
    hf.Close
    Call WriteSuccessMsg("顶部栏目菜单参数设置成功！", ComeUrl)
End Sub

Sub ShowCreate_RootClass_Menu()
    Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong> 生 成 顶 部 栏 目 菜 单 </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150'>"
    Response.Write "<form name='myform' method='post' action='Admin_RootClass_Menu.asp'>"
    Response.Write "<p align='center'>此操作将根据顶部栏目菜单参数设置中设置的参数生成自定义的菜单。</p>"
    Response.Write "<p align='center'><input name='Action' type='hidden' id='Action' value='Create'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='submit' name='Submit' value=' 生成顶部栏目菜单 '></p>"
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
    Call WriteSuccessMsg("顶部栏目菜单生成成功！", ComeUrl)
End Sub

'=================================================
'函数名：GetRootClass_Menu
'作  用：得到栏目无级下拉菜单效果的HTML代码
'参  数：无
'返回值：栏目无级下拉菜单效果的HTML代码
'=================================================
Function GetRootClass_Menu()
    Dim Class_MenuTitle, strJS
    pNum = 1
    pNum2 = 0
    strJS = stm_bm() & vbCrLf
    strJS = strJS & stm_bp_h() & vbCrLf
    strJS = strJS & stm_ai() & vbCrLf
    If UseCreateHTML > 0 Then
        strJS = strJS & stm_aix("p0i1", "p0i0", ChannelName & "首页", ChannelUrl & "/Index" & FileExt_List, "_self", "", False) & vbCrLf
    Else
        strJS = strJS & stm_aix("p0i1", "p0i0", ChannelName & "首页", ChannelUrl & "/Index.asp", "_self", "", False) & vbCrLf
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
    Response.Write stm_aix("p0i1", "p0i0", "频道中心首页", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i3", "p0i0", "学习资料", "#", "_self", "", False) & vbCrLf
    Response.Write stm_bp_v("p1") & vbCrLf
    Response.Write stm_aix("p1i0", "p0i0", "数据库资料", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p1i1", "p0i0", "ASP资料", "#", "_self", "", True) & vbCrLf
    Response.Write stm_bpx("p2", "p1", 1) & vbCrLf
    Response.Write stm_aix("p2i0", "p1i0", "编程技巧", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i1", "p1i0", "经典源码", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i2", "p1i0", "函数合集", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i3", "p1i0", "语法速查", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p2i2", "p1i0", "组件技术", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p2i3", "p1i0", "服务器配置", "#", "_self", "", True) & vbCrLf
    Response.Write stm_bpx("p3", "p2", 1) & vbCrLf
    Response.Write stm_aix("p3i0", "p2i0", "WEB服务器", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i1", "p2i0", "FTP服务器", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p3i4", "p2i0", "网络安全", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i5", "p2i0", "其它资料", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i4", "p0i0", "网络编程", "#", "_self", "", False) & vbCrLf
    Response.Write stm_bpx("p3", "p2", 0) & vbCrLf
    Response.Write stm_aix("p3i0", "p2i0", "PHP编程", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i1", "p2i0", "JSP编程", "#", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p3i2", "p2i0", ".NET编程", "#", "_self", "", False) & vbCrLf
    Response.Write "stm_ep();" & vbCrLf
    Response.Write stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
    Response.Write stm_aix("p0i5", "p0i0", "电子书籍", "#", "_self", "", False) & vbCrLf
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
