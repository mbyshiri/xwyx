<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

If ChannelID = 0 Then
    Response.Write "Ƶ��������ʧ��"
    Call CloseConn
    Response.End
End If
If ModuleType > 5 Then
    Response.Write "<li>ָ����Ƶ��ID���ԣ�</li>"
    Call CloseConn
    Response.End
End If
Dim HtmlDir
HtmlDir = InstallDir & ChannelDir

Response.Write "���ڸ���" & ChannelShortName & "������״̬����"

Dim rsCreate, InfoPath, iCount, iTemp, TheFile, LastModifyTime, NeedUpdate
iCount = PE_Clng(Conn.Execute("select count(0) from PE_" & ModuleName & " where ChannelID=" & ChannelID & " and UpdateTime<CreateTime")(0))
Response.Write "һ����Ҫ��� " & iCount & " " & ChannelItemUnit & "���ݿ��б�ʶΪ�������ɡ���" & ChannelShortName & "��"
iCount = PE_Clng(Conn.Execute("select count(0) from PE_" & ModuleName & " where ChannelID=" & ChannelID & " and (CreateTime is null or CreateTime<=UpdateTime)")(0))
Response.Write "�������� " & iCount & " " & ChannelItemUnit & ChannelShortName & "��ʶΪ��δ���ɡ�������Ҫ�������״̬��<br>"

iCount = 0
iTemp = 0
Set rsCreate = Conn.Execute("select C.ParentDir,C.ClassDir,I." & ModuleName & "ID as InfoID,I.UpdateTime,I.CreateTime from PE_" & ModuleName & " I  left join PE_Class C on I.ClassID=C.ClassID where I.ChannelID=" & ChannelID & " and UpdateTime<CreateTime")
Do While Not rsCreate.EOF
    NeedUpdate = False
    InfoPath = HtmlDir & GetItemPath(StructureType, rsCreate("ParentDir"), rsCreate("ClassDir"), rsCreate("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsCreate("UpdateTime"), rsCreate("InfoID")) & FileExt_Item
    If fso.FileExists(Server.MapPath(InfoPath)) Then
        If Not IsDate(rsCreate("CreateTime")) Then
            NeedUpdate = True
        Else
            Set TheFile = fso.GetFile(Server.MapPath(InfoPath))
            LastModifyTime = TheFile.DateLastModified
            If rsCreate("UpdateTime") > LastModifyTime Then
                NeedUpdate = True
            End If
        End If
    Else
        NeedUpdate = True
    End If
    If NeedUpdate = True Then
        Conn.Execute ("update PE_" & ModuleName & " set CreateTime=UpdateTime where " & ModuleName & "ID=" & rsCreate("InfoID") & "")
        iCount = iCount + 1
    End If
    iTemp = iTemp + 1
    If iTemp Mod 10 = 0 Then
        Response.Write "."
        Response.Flush
    End If
    If iTemp Mod 1000 = 0 Then
        Response.Write "<br>"
        Response.Flush
    End If
    rsCreate.MoveNext
Loop
rsCreate.Close
Set rsCreate = Nothing
Call CloseConn
Response.Write "<br><br>����" & ChannelShortName & "������״̬��ɣ�"
Response.Write "��鷢�ֹ��� " & iCount & " " & ChannelItemUnit & ChannelShortName & "ʵ������δ���ɵģ��Ѿ�����������״̬��"
Response.Write "<p align='center'><a href='Admin_CreateHTML.asp?ChannelID=" & ChannelID & "'>�����ء�</a></p>" & vbCrLf
%>
