<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Cache"   '其他权限

Dim CacheName, ShowType, ShowTypeName
CacheName = "PowerEasy"
ShowType = Trim(Request("ShowType"))
If ShowType = "" Then
    ShowType = 1
Else
    ShowType = PE_CLng(ShowType)
End If
If ShowType = 1 Then
    ShowTypeName = "缓存"
ElseIf ShowType = 2 Then
    ShowTypeName = "变量"
End If

'页面头部HTML代码
Response.Write "<html><head><title>网站缓存管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <tr class='topbg'>" & vbCrLf
Response.Write "    <td height='22' colspan='2' align='center'><strong>网 站 缓 存 管 理</strong></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td>" & vbCrLf
Response.Write "      <a href='Admin_Cache.asp?ShowType=1'>网站缓存管理</a>&nbsp;|&nbsp;"
If AdminPurview = 1 Then
    Response.Write "      <a href='Admin_Cache.asp?ShowType=2'>服务器Application变量</a>&nbsp;|&nbsp;"
End If
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

'执行的操作
Select Case Action
Case "Show"
    Call Show
Case "Del"
    Call Del
Case "Clear"
    Call Clear
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetManagePath & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Cache.asp'>"
    Response.Write "    <td>"
    Response.Write "      <table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "        <tr class='title' height='22'> "
    Response.Write "          <td width='200' align='center'><strong>" & ShowTypeName & "名称</strong></td>"
    Response.Write "          <td align='center'><strong>" & ShowTypeName & "值</strong></td>"
    If ShowType = 1 Then
        Response.Write "          <td width='150' align='center'><strong>" & ShowTypeName & "时间</strong></td>"
    End If
    Response.Write "          <td width='100' align='center'><strong>操作</strong></td>"
    Response.Write "        </tr>"

    If Application.Contents.Count = 0 Then
        Response.Write "        <tr class='tdbg'><td colspan='20' align='center'><br>没有任何缓存！<br><br></td></tr>"
    Else
        Dim Item, CacheObj, ShowFlag
        Set CacheObj = Application.Contents
        For Each Item In CacheObj
            ShowFlag = False
            If ShowType = 1 Then
                If CStr(Left(Item, Len(CacheName) + 1)) = CStr(CacheName & "_") Then
                    ShowFlag = True
                End If
            Else
                If AdminPurview = 1 And CStr(Left(Item, Len(CacheName) + 1)) <> CStr(CacheName & "_") And InStr(LCase(CStr(Item)), "conn") = 0 And InStr(LCase(CStr(Item)), "dbpath") = 0 And InStr(LCase(CStr(Item)), "sitekey") = 0 Then
                    ShowFlag = True
                End If
            End If
            If ShowFlag = True Then
                Response.Write "        <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
                Response.Write "          <td width='200' align='center'>"
                Response.Write "            <a href='Admin_Cache.asp?Action=Show&Name=" & Item & "'>" & Item & "</a>"
                Response.Write "          </td>"
                Response.Write "          <td>"
                Response.Write Left(nohtml(PE_HTMLEncode(GetApplication(CacheObj(Item)))), 90)
                Response.Write "          </td>"
                If ShowType = 1 Then
                    Response.Write "          <td width='150' align='center'>"
                    Response.Write GetAppTime(CacheObj(Item))
                    Response.Write "          </td>"
                End If
                Response.Write "          <td width='100' align='center'>"
                Response.Write "            <a href='Admin_Cache.asp?Action=Show&Name=" & Item & "'>查看</a> | "
                Response.Write "            <a href='Admin_Cache.asp?Action=Del&Name=" & Item & "' onClick=""return confirm('确定要删除此" & ShowTypeName & "吗？');"">删除</a>"
                Response.Write "          </td>"
                Response.Write "        </tr>"
            End If
        Next
    End If
    Response.Write "      </table>"
    If ShowType = 1 Then
        Response.Write "      <table width='100%'><form name='form1' action='Admin_Cache.asp' method='post'>"
        Response.Write "        <tr>"
        Response.Write "          <td align='center'>"
        Response.Write "            <input name='Action' type='hidden' id='Action' value='Clear'><input type='submit' value='清除所有缓存' name='submit'>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "      </form></table>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"
End Sub

Function GetManagePath()
    Dim strPath
    strPath = "您现在的位置："
    If ShowType = 1 Then
        strPath = strPath & "网站缓存管理"
    Else
        strPath = strPath & "服务器Application变量"
    End If
    GetManagePath = strPath
End Function

Sub Show()
    Dim ApplicationName
    ApplicationName = Trim(Request("Name"))
    If ApplicationName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定缓存</li>"
        Exit Sub
    End If
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>您现在的位置：查看" & ShowTypeName & "值</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "  <tr class='title' height='22'> "
    Response.Write "    <td align='center'><strong>" & ApplicationName & " 值</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='400' align='center'><textarea cols='110' rows='25'>" & GetApplication(Application(ApplicationName)) & "</textarea></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Function GetApplication(CacheObjItem)
    On Error Resume Next
    Dim strApplication
    If IsArray(CacheObjItem) Then
        strApplication = strApplication & CacheObjItem(0)
    Else
        strApplication = strApplication & CacheObjItem
    End If
    GetApplication = strApplication
End Function

Function GetAppTime(CacheObjItem)
    On Error Resume Next
    Dim strAppTime
    If IsArray(CacheObjItem) Then
        If UBound(CacheObjItem) > 0 Then
            strAppTime = strAppTime & CacheObjItem(1)
        End If
    End If
    GetAppTime = strAppTime
End Function

Sub Del()
    Dim ApplicationName
    ApplicationName = Trim(Request("Name"))
    If ApplicationName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定缓存</li>"
        Exit Sub
    End If
    Application.Lock
    Application.Contents.Remove ApplicationName
    Application.UnLock
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub Clear()
    Call PE_Cache.DelAllCache
    Call WriteSuccessMsg("清除所有缓存成功。", ComeUrl)
End Sub
%>
