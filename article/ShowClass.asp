<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

ClassID = PE_CLng(Trim(Request("ClassID")))
If ClassID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ָ����ĿID��</li>"
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
        ErrMsg = ErrMsg & XmlText("BaseText", "PurviewCheckedErr", "<li>�Բ�����û���������Ŀ���ݵ�Ȩ�ޣ�</li>")
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
    Response.Write "��ǰ��Ŀʹ�õ�ҳ��ģ������ȱ��С��ģ�壡"
    Response.End
End If
Call PE_Content.GetHtml_Class

Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
