<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ClassID = PE_CLng(Trim(Request("ClassID")))
If ClassID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>请指定栏目ID！</li>"
Else
    Call GetClass
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

ItemID = 0
If ClassPurview > 1 Then
    Dim PurviewChecked
    If UserLogined = False Then
        PurviewChecked = False
    Else
        Call GetUser(UserName)
        If ParentID > 0 Then
            PurviewChecked = CheckPurview_Class(arrClass_Browse, ChannelDir & "all," & ParentPath & "," & ClassID)
        Else
            PurviewChecked = CheckPurview_Class(arrClass_Browse, ChannelDir & "all," & ClassID)
        End If
    End If
    If PurviewChecked = False Then
        FoundErr = True
        ErrMsg = ErrMsg & XmlText("BaseText", "PurviewCheckedErr", "<li>对不起，您没有浏览此栏目内容的权限！</li>")
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Response.End
    End If
End If
ClassShowType = Trim(Request("ShowType"))
If ClassShowType = "" Then
    ClassShowType = 1
Else
    ClassShowType = PE_CLng(ClassShowType)
End If

PageTitle = ""
If ClassShowType = 2 Then
    strFileName = ChannelUrl_ASPFile & "/ShowClass.asp?ShowType=2&ClassID=" & ClassID
Else
    strFileName = ChannelUrl_ASPFile & "/ShowClass.asp?ClassID=" & ClassID
End If
strTemplate = GetTemplate(ChannelID, 2, TemplateID)
arrTemplate = Split(strTemplate, "{$$$}")
If UBound(arrTemplate) < 1 Then
    Response.Write "当前栏目使用的页面模板有误，缺少小类模板！"
    Response.End
End If
Call PE_Content.GetHtml_Class

Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
