<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

ArticleID = PE_CLng(Trim(Request("ArticleID")))
If ArticleID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ָ��ArticleID��</li>"
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
Dim sql
PageTitle = "����"
ItemID = ArticleID
strFileName = "ShowArticle.asp"

Conn.Execute ("update PE_Article set Hits=Hits+1 where ArticleID=" & ArticleID)
    
sql = "select * from PE_Article where Deleted=" & PE_False & " and Status=3 and ArticleID=" & ArticleID & " and ChannelID=" & ChannelID & ""
Set rsArticle = Conn.Execute(sql)
If rsArticle.BOF And rsArticle.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��Ҫ�ҵ�" & ChannelShortName & "�����ڣ������Ѿ�������Աɾ����</li>"
Else
    ClassID = rsArticle("ClassID")
    If ClassID > 0 Then
        Call GetClass
    End If
    If rsArticle("ReceiveType") = 1 Then
        If UserLogined = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ȩ�����ר��" & ChannelShortName & "��</li>"
        Else
            If FoundInArr(rsArticle("ReceiveUser"), UserName, ",") = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>����Ȩ�����ר��" & ChannelShortName & "��</li>"
            End If
        End If
    End If
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    rsArticle.Close
    Set rsArticle = Nothing
    Response.End
End If

If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then
    Response.Write "<script language='javascript'>window.location.href='" & rsArticle("LinkUrl") & "';</script>"
Else
    If Trim(rsArticle("TitleIntact")) <> "" Then
        ArticleTitle = Replace(Replace(Replace(Replace(rsArticle("TitleIntact") & "", "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    Else
        ArticleTitle = Replace(Replace(Replace(Replace(rsArticle("Title") & "", "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    End If

    SkinID = GetIDByDefault(rsArticle("SkinID"), DefaultItemSkin)
    TemplateID = GetIDByDefault(rsArticle("TemplateID"), DefaultItemTemplate)

    ArticleUrl = GetArticleUrl(ParentDir, ClassDir, rsArticle("UpdateTime"), ArticleID, ClassPurview, rsArticle("InfoPurview"), rsArticle("InfoPoint"))
    strHtml = GetTemplate(ChannelID, 3, TemplateID)
    Call PE_Content.GetHtml_Article
    strHtml = PE_Content.ReplaceContentLabel(strHtml)
    Response.Write strHtml
End If
rsArticle.Close
Set rsArticle = Nothing
Set PE_Content = Nothing
Call CloseConn
%>
