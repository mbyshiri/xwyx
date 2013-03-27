<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

If ChannelID = 0 Then
    Response.Write "频道参数丢失！"
    Call CloseConn
    Response.End
End If
If ModuleType > 5 Then
    Response.Write "<li>指定的频道ID不对！</li>"
    Call CloseConn
    Response.End
End If
Dim HtmlDir
HtmlDir = InstallDir & ChannelDir

Response.Write "正在更新" & ChannelShortName & "的生成状态……"

Dim rsCreate, InfoPath, iCount, iTemp, TheFile, LastModifyTime, NeedUpdate
iCount = PE_Clng(Conn.Execute("select count(0) from PE_" & ModuleName & " where ChannelID=" & ChannelID & " and UpdateTime<CreateTime")(0))
Response.Write "一共需要检查 " & iCount & " " & ChannelItemUnit & "数据库中标识为“已生成”的" & ChannelShortName & "，"
iCount = PE_Clng(Conn.Execute("select count(0) from PE_" & ModuleName & " where ChannelID=" & ChannelID & " and (CreateTime is null or CreateTime<=UpdateTime)")(0))
Response.Write "其他还有 " & iCount & " " & ChannelItemUnit & ChannelShortName & "标识为“未生成”，不需要检查生成状态。<br>"

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
Response.Write "<br><br>更新" & ChannelShortName & "的生成状态完成！"
Response.Write "检查发现共有 " & iCount & " " & ChannelItemUnit & ChannelShortName & "实际上是未生成的，已经更新其生成状态。"
Response.Write "<p align='center'><a href='Admin_CreateHTML.asp?ChannelID=" & ChannelID & "'>【返回】</a></p>" & vbCrLf
%>
