<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = False   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim resultMsg
Dim CheckType
CheckType = Trim(Request("CheckType"))
If CheckType = "CheckNum" Then
    Call JudgeSameNum
Else
    Call JudgeSameTitle
End If
Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>检测重复" & ModuleName & "</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<base target='_self'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<Table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border' Height='100%'>" & vbCrLf
Response.Write "<tr class='title' height='22'>" & vbCrLf
Response.Write "    <td align=center>" & vbCrLf
Response.Write "    <b>检测结果<b>" & vbCrLf
Response.Write "    </td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr class='tdbg'>" & vbCrLf
Response.Write "    <td valign='top' align='left'>" & vbCrLf
Response.Write resultMsg
Response.Write "    </td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr class='title' height='22'>" & vbCrLf
Response.Write "    <td align=center>" & vbCrLf
Response.Write "        <input type=button name='button1' id='button1' value='关闭窗口' onClick='javascript:window.close();'>" & vbCrLf
Response.Write "    </td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "</Table>"
Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf

Sub JudgeSameNum()
    Dim strSql, rsTitle, Title
    Title = Trim(Request("Title"))
    If Title = "" Then
        resultMsg = "被检测商品的编号不可以为空！"
        Exit Sub
    End If

    Title = ReplaceText(Title, 2)
    Title = PE_HTMLEncode(Title)
    strSql = "Select ProductNum,ProductName From PE_Product Where ProductNum='" & Title & "'"
    Set rsTitle = Server.CreateObject("Adodb.Recordset")
    rsTitle.Open strSql, Conn, 1, 1
    If rsTitle.EOF And rsTitle.BOF Then
        resultMsg = "该" & ModuleName & "编号没有被使用！"
    Else
        resultMsg = "此" & ModuleName & "编号已存在，相关如下:<br>"
        Do While Not rsTitle.EOF
            resultMsg = resultMsg & "<li>编号：" & rsTitle(0) & "&nbsp;&nbsp;商品名:" & rsTitle(1) & "</li><br>"
            rsTitle.movenext
        Loop
    End If
    rsTitle.Close
    Set rsTitle = Nothing
End Sub
'判断标题是否重复
Sub JudgeSameTitle()
    Dim strSql, Title, TableName, ModuleType
    Dim rsTitle
    Title = Trim(Request("Title"))
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
    
    If Title = "" Then
        resultMsg = "被检测文章的标题不可以为空！"
        Exit Sub
    End If

    Title = ReplaceText(Title, 2)
    Title = PE_HTMLEncode(Title)
    Select Case ModuleType
        Case "1"
            strSql = "Select  Title  From PE_Article Where Title='" & Title & "'"
            ModuleName = "标题"
        Case "2"
            strSql = "Select SoftName From PE_Soft Where SoftName = '" & Title & "'"
            ModuleName = "软件名"
        Case "3"
            strSql = "Select PhotoName From PE_Photo Where PhotoName='" & Title & "'"
            ModuleName = "图片名"
        Case "5"
            strSql = "Select ProductName From PE_Product Where ProductName='" & Title & "'"
            ModuleName = "商品名"
        Case Else
            resultMsg = "错误的传递参数！"
            Exit Sub
    End Select
    
    Set rsTitle = Server.CreateObject("Adodb.Recordset")
    rsTitle.Open strSql, Conn, 1, 1
    If rsTitle.EOF And rsTitle.BOF Then
        resultMsg = "该" & ModuleName & "没有被使用！"
    Else
        resultMsg = "此" & ModuleName & "已存在，相关如下:<br>"
        Do While Not rsTitle.EOF
            resultMsg = resultMsg & "<li>" & rsTitle(0) & "</li><br>"
            rsTitle.movenext
        Loop
    End If
    rsTitle.Close
    Set rsTitle = Nothing
End Sub
%>
