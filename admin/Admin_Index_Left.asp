<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Edition.asp"-->
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

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<title>�������˵�</title>" & vbCrLf
Response.Write "<script src=""../JS/prototype.js""></script>"
Response.Write "<link href='Admin_left.CSS' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<BODY leftmargin='0' topmargin='0' marginheight='0' marginwidth='0'>" & vbCrLf
Response.Write "<table width=180 border='0' align=center cellpadding=0 cellspacing=0>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=44 valign=top><img src='Images/title.gif'></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "<table cellpadding=0 cellspacing=0 width=180 align=center>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=26 class=menu_title onmouseover=""this.className='menu_title2';"" onmouseout=""this.className='menu_title';"" background='Images/title_bg_quit.gif' id='menuTitle0'>&nbsp;&nbsp;<a href='Admin_Index_Main.asp' target='main'><b><span class='glow'>������ҳ</span></b></a><span class='glow'> | </span><a href='Admin_Login.asp?Action=Logout' target='_top'><b><span class='glow'>�˳�</span></b></a> </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=97 background='Images/title_bg_admin.gif' style='display:' id='submenu0'><div style='width:180'>" & vbCrLf
Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
Response.Write "          <tr>" & vbCrLf
Response.Write "            <td height=25>�����û�����" & AdminName & "</td>" & vbCrLf
Response.Write "          </tr>" & vbCrLf
Response.Write "          <tr>" & vbCrLf
Response.Write "            <td height=25>������ݣ�"
Select Case AdminPurview
Case 1
    Response.Write "��������Ա"
Case 2
    Response.Write "<a href='Admin_ShowPurview.asp' target='main'>��ͨ����Ա</a>"
End Select
Dim Message
Set Message = Conn.Execute("select Count(0) from PE_Message where Incept = '" & UserName & "' and delR=0 and Flag=0 and IsSend=1")
If Message.EOF And Message.Bof Then
    UnreadMsg = 0
Else
    UnreadMsg = Message(0)
End If
Set Message = Nothing
Response.Write "</td></tr><tr><td height=20>���Ķ��ţ�" & vbCrLf
If UnreadMsg > 0 Then
    Response.Write " <b><font color=red>" & UnreadMsg & "</font></b> ��"
Else
    Response.Write " <b><font color=gray>0</font></b> ��"
End If
Response.Write "            </td>" & vbCrLf
Response.Write "          </tr>" & vbCrLf
Response.Write "        </table>" & vbCrLf
Response.Write "      </div>" & vbCrLf
Response.Write "        <div  style='width:167'>" & vbCrLf
Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
Response.Write "            <tr>" & vbCrLf
Response.Write "              <td height=20></td>" & vbCrLf
Response.Write "            </tr>" & vbCrLf
Response.Write "          </table>" & vbCrLf
Response.Write "      </div></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Dim ShowCreateHTML, strActionLink, ShowAdmin_Guest, ShowAdmin_Shop
ShowAdmin_Shop = False
ShowCreateHTML = False

Dim sqlChannel, rsChannel
sqlChannel = "select * from PE_Channel where ChannelType<=1"
Select Case SystemEdition
Case "CMS", "GPS", "EPS"
    sqlChannel = sqlChannel & " and ModuleType<5"
Case "eShop", "ECS"
    sqlChannel = sqlChannel & " and ModuleType<6"
Case "IPS", "All"
    sqlChannel = sqlChannel & " and ModuleType<7"
End Select
sqlChannel = sqlChannel & " order by OrderID"
Set rsChannel = Conn.Execute(sqlChannel)
Do While Not rsChannel.EOF
    If rsChannel("ModuleType") = 4 Then
        If rsChannel("Disabled") = True Then
            ShowAdmin_Guest = False
        Else
            ShowAdmin_Guest = True
        End If
    Else
        If rsChannel("Disabled") = False Then
            ChannelID = rsChannel("ChannelID")
            ChannelName = Trim(rsChannel("ChannelName"))
            ChannelShortName = Trim(rsChannel("ChannelShortName"))
            ChannelDir = Trim(rsChannel("ChannelDir"))
            Select Case rsChannel("ModuleType")
            Case 1
                ModuleName = "Article"
            Case 2
                ModuleName = "Soft"
            Case 3
                ModuleName = "Photo"
            Case 5
                ModuleName = "Product"
            Case 6
                ModuleName = "Supply"
            End Select
            AdminPurview_Channel = rsGetAdmin("AdminPurview_" & ChannelDir)
            If IsNull(AdminPurview_Channel) Then
                AdminPurview_Channel = 5
            Else
                AdminPurview_Channel = CLng(AdminPurview_Channel)
            End If

            
            If AdminPurview = 1 Or AdminPurview_Channel <= 3 Then
                Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center><tr>" & vbCrLf
                Response.Write "<td height=28 class=menu_title onmouseover=""this.className='menu_title2'""; onmouseout=""this.className='menu_title'""; background='Images/Admin_left_" & rsChannel("ModuleType") & ".gif' id=menuTitle" & ChannelID & " onclick=""new Element.toggle('submenu" & ChannelID & "')"" style='cursor:hand;'>" & vbCrLf
                If rsChannel("ModuleType") = 6 Then
                    Response.Write "<a href='Admin_Help_Supply.asp?ChannelID=" & ChannelID & "' target=main><span class=glow>" & ChannelName & "����</span></a>"
                Else
                    Response.Write "<a href='Admin_Help_Channel.asp?ChannelID=" & ChannelID & "' target=main><span class=glow>" & ChannelName & "����</span></a>"
                End If
                Response.Write "</td></tr><tr><td style='display:none' align='right' id='submenu" & ChannelID & "'><div class=sec_menu style='width:165'><table cellpadding=0 cellspacing=0 align=center width=132>" & vbCrLf
                If rsChannel("ModuleType") = 5 Then
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&ProductKind=1' target=main>���" & ChannelShortName & "��ʵ�</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&ProductKind=2' target=main>���" & ChannelShortName & "�������</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&ProductKind=3' target=main>���" & ChannelShortName & "���㿨��</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_Card.asp' target=main>��ֵ������</a></td></tr>" & vbCrLf
                Else
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&AddType=1' target=main>���" & ChannelShortName & "</a></td></tr>" & vbCrLf
                    If rsChannel("ModuleType") = 1 Then
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Receive&Status=9' target=main>ǩ��" & ChannelShortName & "����</a></td></tr>"
                    End If
                    If rsChannel("ModuleType") = 2 Then
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&AddType=3' target=main>���" & ChannelShortName & "������ģʽ��</a></td></tr>" & vbCrLf
                    End If
                    If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                        If rsChannel("ModuleType") = 3 Then
                            Response.Write "<tr><td height=20><a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=3' target=main>���" & ChannelShortName & "������ģʽ��</a></td></tr>" & vbCrLf
                        End If
                    End If
                End If
                Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=My&Status=9' target=main>����ӵ�" & ChannelShortName & "</a></td></tr>" & vbCrLf
                Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage' target=main>" & ChannelShortName & "����</a>"
                If rsChannel("ModuleType") = 5 Then
                    Response.Write " | <a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Price&Status=0' target=main>�۸�����</a>"
                Else
                    Response.Write " | <a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Check&Status=0' target=main>���</a>"
                End If
                If rsChannel("UseCreateHTML") > 0 And ObjInstalled_FSO = True Then
                    ShowCreateHTML = True
                    Response.Write " | <a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&Status=3&ManageType=HTML' target=main>����</a>" & vbCrLf
                    If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
                        strActionLink = strActionLink & "<tr height='20'><td><a href='Admin_CreateHTML.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelName & "���ɹ���</a></td></tr>"
                    End If
                End If
                Response.Write "</td></tr>"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Special&Status=9' target=main>ר��" & ChannelShortName & "����</a></td></tr>" & vbCrLf
                End If
                If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
                    If rsChannel("ModuleType") = 2 Then
                        Response.Write "<tr><td height=20><a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=ShowReplace' target=main>���ص�ַ�����޸�</a></td></tr>" & vbCrLf
                    End If
                    Response.Write "<tr><td height=20><a href='Admin_Comment.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelShortName & "���۹���</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Recyclebin' target=main>" & ChannelShortName & "����վ����</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=10>=====================</td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_Channel.asp?Action=Modify&iChannelID=" & ChannelID & "' target=main>" & ChannelName & "����</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_Class.asp?ChannelID=" & ChannelID & "' target=main>��Ŀ����</a> | <a href='Admin_Special.asp?ChannelID=" & ChannelID & "' target=main>ר�����</a></td></tr>" & vbCrLf
                    Select Case rsChannel("ModuleType")
                    Case 1, 5
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadFiles' target=main>�ϴ��ļ�����</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadFiles' target=main>����</a></td></tr>" & vbCrLf
                    Case 2
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadSoftPic' target=main>�ϴ�ͼƬ����</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadSoftPic' target=main>����</a></td></tr>" & vbCrLf
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadSoft' target=main>�ϴ��ļ�����</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadSoft' target=main>����</a></td></tr>" & vbCrLf
                    Case 3, 6
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadPhotos' target=main>�ϴ�ͼƬ����</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadPhotos' target=main>����</a></td></tr>" & vbCrLf
                    End Select

                    If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Menu_" & ChannelDir) = True Then
                        Response.Write "<tr><td height=20><a href='Admin_RootClass_Menu.asp?ChannelID=" & ChannelID & "' target=main>�����˵�����</a> | <a href='Admin_RootClass_Menu.asp?Action=ShowCreate&ChannelID=" & ChannelID & "' target=main>����</a></td></tr>" & vbCrLf
                    End If
                    If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir) = True Or CheckPurview_Other(AdminPurview_Others, "JsFile_" & ChannelDir) = True Then
                        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir) = True Then
                            Response.Write "<tr><td height=20><a href='Admin_Template.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelShortName & "ģ��ҳ����</a></td></tr>" & vbCrLf
                        End If
                        If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "JsFile_" & ChannelDir) = True And ObjInstalled_FSO = True) And (rsChannel("ModuleType") <> 6) Then
                            Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelShortName & "JS�ļ�����</a></td></tr>" & vbCrLf
                        End If
                    End If
                    If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Keyword_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                        Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword' target='main'>" & ChannelShortName & "�ؼ��ֹ���</a></td></tr>"
                    End If
                    If rsChannel("ModuleType") = 5 Then
                        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Producer_" & ChannelDir) = True Then
                            Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer' target='main'>���̹���</a></td></tr>"
                        End If
                        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Trademark_" & ChannelDir) = True Then
                            Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark' target='main'>Ʒ�ƹ���</a></td></tr>"
                        End If
                    Else
                        If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Author_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                            Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author' target='main'>���߹���</a>"
                        End If
                        If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Copyfrom_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                            Response.Write " | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom' target='main'>��Դ����</a></td></tr>"
                        End If
                    End If
                    If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "XML_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                        Response.Write "<tr><td height=20><a href='Admin_CreateXml.asp?Action=GreatNav&ChannelID=" & ChannelID & "' target=main>������ĿXML����</a></td></tr>" & vbCrLf
                    End If
                    If rsChannel("ModuleType") = 2 Then
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Other' target=main>��������</a></td></tr>" & vbCrLf
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=DownError' target=main>������Ϣ����</a></td></tr>" & vbCrLf
                        Response.Write "<tr><td height=20><a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "' target=main>�������������</a></td></tr>" & vbCrLf
                    End If
                    If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Field_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                        Response.Write "<tr><td height=20><a href='Admin_Field.asp?ChannelID=" & ChannelID & "' target=main>�Զ����ֶι���</a></td></tr>" & vbCrLf
                    End If
                    If rsChannel("ModuleType") = 1 Then
                        If FoundInArr(rsChannel("arrEnabledTabs"), "Copyfee", ",") = True Then
                            Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ManageType=PayMoney&PayStatus=False' target=main>��ѹ���</a></td></tr>"
                        End If
                    End If
                End If
                Response.Write "</table></div><div  style='width:158'><table cellpadding=0 cellspacing=0 align=center width=130><tr><td height=4></td></tr></table></div></td></tr></table>" & vbCrLf
            End If
            If rsChannel("ModuleType") = 5 Then ShowAdmin_Shop = True
        End If
    End If
    rsChannel.MoveNext
Loop
rsChannel.Close
Set rsChannel = Nothing

If (SystemEdition = "IPS" Or SystemEdition = "All") And FoundInArr(AllModules, "House", ",") = True Then
    Dim rsHouse, rsHouseConfig
    Set rsHouse = Conn.Execute("Select ChannelID, ChannelDir, ChannelName, Disabled from PE_Channel Where ModuleType=7")
    If rsHouse("Disabled") = False Then
        ChannelDir = rsHouse("ChannelDir")
        AdminPurview_Channel = rsGetAdmin("AdminPurview_" & ChannelDir)
        If AdminPurview = 1 Or AdminPurview_Channel <= 3 Then
            Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_7.gif' id=menuTitle502 onclick=""new Element.toggle('submenu502')"" style='cursor:hand;'><a href='Admin_Help_House.asp' target='main'><span>" & rsHouse("ChannelName") & "����</span></a></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td style='display:none' align='right' id='submenu502'><div class=sec_menu style='width:165'>" & vbCrLf
            Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Set rsHouseConfig = Conn.Execute("Select ClassID,ClassName from PE_HouseConfig")
            Do While Not rsHouseConfig.EOF
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Add&ClassID=" & rsHouseConfig("ClassID") & "' target='main'>����" & rsHouseConfig("ClassName") & "��Ϣ</a> | <a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Manage&ClassID=" & rsHouseConfig("ClassID") & "' target='main'>����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                rsHouseConfig.MoveNext
            Loop
            rsHouseConfig.Close
            Set rsHouseConfig = Nothing
            Response.Write "<tr><td height=10>=====================</td></tr>" & vbCrLf
            If AdminPurview = 1 Or arrPurview(2) = True Or CheckPurview_Other(AdminPurview_Others, "Area_" & ChannelDir) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=ManageArea' target='main'>�������</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or arrPurview(2) = True Or CheckPurview_Other(AdminPurview_Others, "ClassConfig_" & ChannelDir) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=ModifyConfig' target='main'>������Ŀ����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & rsHouse("ChannelID") & "&UploadDir=UploadPhotos' target=main>�ϴ�ͼƬ����</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Clear&UploadDir=UploadPhotos' target=main>����</a></td></tr>" & vbCrLf
            If AdminPurview = 1 Or arrPurview(2) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Manage&ManageType=RecycleBin' target='main'>����վ����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or arrPurview(2) = True Or CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=" & rsHouse("ChannelID") & "' target='main'>ģ�����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "        </table>" & vbCrLf
            Response.Write "      </div>" & vbCrLf
            Response.Write "        <div  style='width:158'>" & vbCrLf
            Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Response.Write "            <tr>" & vbCrLf
            Response.Write "              <td height=5></td>" & vbCrLf
            Response.Write "            </tr>" & vbCrLf
            Response.Write "          </table>" & vbCrLf
            Response.Write "      </div></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
        End If
    End If
    rsHouse.Close
    Set rsHouse = Nothing
End If
If (SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "All") And FoundInArr(AllModules, "Job", ",") = True Then
    Dim rsJob
    Set rsJob = Conn.Execute("Select ChannelName, ChannelDir, Disabled from PE_Channel Where ModuleType=8")
    If rsJob("Disabled") = False Then
        ChannelDir = rsJob("ChannelDir")
        AdminPurview_Channel = rsGetAdmin("AdminPurview_" & ChannelDir)
        If AdminPurview = 1 Or AdminPurview_Channel <= 3 Then
            Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_8.gif' id=menuTitle607 onclick=""new Element.toggle('submenu607')"" style='cursor:hand;'><a href='Admin_Help_Job.asp' target='main'><span>" & rsJob("ChannelName") & "����</span></a></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td style='display:none' align='right' id='submenu607'><div class=sec_menu style='width:165'>" & vbCrLf
            Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=Position' target='main'>ְλ��Ϣ����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or AdminPurview_Channel = 1 Or AdminPurview_Channel = 3 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=Resume' target='main'>�˲���Ϣ����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "          <tr height=10>" & vbCrLf
            Response.Write "            <td>====================</td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=JobCategory' target='main'>����������</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=WorkPlace' target='main'>�����ص����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=SubCompany' target='main'>���˵�λ����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_UploadFile.asp?ChannelID=997&UploadDir=UploadPhotos' target=main>�ϴ�ͼƬ����</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=997&Action=Clear&UploadDir=UploadPhotos' target=main>����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=997' target='main'>��Ƹģ��ҳ����</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "        </table>" & vbCrLf
            Response.Write "      </div>" & vbCrLf
            Response.Write "        <div  style='width:158'>" & vbCrLf
            Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Response.Write "            <tr>" & vbCrLf
            Response.Write "              <td height=5></td>" & vbCrLf
            Response.Write "            </tr>" & vbCrLf
            Response.Write "          </table>" & vbCrLf
            Response.Write "      </div></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
        End If
    End If
    rsJob.Close
    Set rsJob = Nothing
End If

If ShowAdmin_Shop = True Then
    Dim Purview_OrderManage
    PurviewPassed = False
    Purview_OrderManage = False
    arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "Order_View")
    arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "Order_Confirm")
    arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "Order_Modify")
    arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "Order_Del")
    arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "Order_Payment")
    arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "Order_Invoice")
    arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "Order_Deliver")
    arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "Order_Download")
    arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "Order_SendCard")
    arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "Order_End")
    arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Order_Transfer")
    arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "Order_Print")
    arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "Order_Count")
    arrPurview(13) = CheckPurview_Other(AdminPurview_Others, "Order_OrderItem")
    arrPurview(14) = CheckPurview_Other(AdminPurview_Others, "Order_SaleCount")
    arrPurview(15) = CheckPurview_Other(AdminPurview_Others, "Payment")
    arrPurview(16) = CheckPurview_Other(AdminPurview_Others, "Bankroll")
    arrPurview(17) = CheckPurview_Other(AdminPurview_Others, "Deliver")
    arrPurview(18) = CheckPurview_Other(AdminPurview_Others, "Transfer")
    arrPurview(19) = CheckPurview_Other(AdminPurview_Others, "PresentProject")
    arrPurview(20) = CheckPurview_Other(AdminPurview_Others, "PaymentType")
    arrPurview(21) = CheckPurview_Other(AdminPurview_Others, "DeliverType")
    arrPurview(22) = CheckPurview_Other(AdminPurview_Others, "Bank")
    arrPurview(23) = CheckPurview_Other(AdminPurview_Others, "ShoppingCart")
    arrPurview(24) = CheckPurview_Other(AdminPurview_Others, "Order_Refund")
    arrPurview(25) = CheckPurview_Other(AdminPurview_Others, "AddPayment")
    arrPurview(26) = CheckPurview_Other(AdminPurview_Others, "AgentPayment")
    For PurviewIndex = 0 To 12
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            Purview_OrderManage = True
            Exit For
        End If
    Next
    For PurviewIndex = 13 To 23
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            Exit For
        End If
    Next
    If arrPurview(24) = True Or arrPurview(25) = True Or arrPurview(26) = True Then
        PurviewPassed = True
        Purview_OrderManage = True
    End If
    If AdminPurview = 1 Or PurviewPassed = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_10.gif' id=menuTitle901 onclick=""new Element.toggle('submenu901')"" style='cursor:hand;'><a href='admin_help_Shop.asp' target='main'><span class=glow>�̳��ճ�����</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu901'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    
        If AdminPurview = 1 Or Purview_OrderManage = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Order.asp?SearchType=1' target='main'>�������Ķ���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Order.asp' target='main'>�������ж���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr height=10>" & vbCrLf
            Response.Write "            <td>====================</td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(13) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_OrderItem.asp' target='main'>������ϸ���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(14) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_SaleCount.asp' target='main'>����ͳ��/����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(15) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Payment.asp' target='main'>����֧����¼����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(16) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Bankroll.asp' target='main'>�ʽ���ϸ��ѯ</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(5) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Invoice.asp' target='main'>����Ʊ��¼</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(17) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Deliver.asp' target='main'>���˻���¼</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(18) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Transfer.asp' target='main'>����������¼</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(23) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_ShoppingCart.asp' target='main'>���ﳵ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(19) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_PresentProject.asp' target='main'>������������</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(20) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_PaymentType.asp' target='main'>���ʽ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(21) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_DeliverType.asp' target='main'>�ͻ���ʽ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If ShowAdmin_Shop = True And FoundInArr(AllModules, "CRM", ",") Then
    Dim PurviewPassed_Client, PurviewPassed_Service, PurviewPassed_Complain, PurviewPassed_Call
    PurviewPassed = False
    arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "Client_View")
    arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "Client_Add")
    arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "Client_ModifyOwn")
    arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "Client_ModifyAll")
    arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "Client_Del")
    arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "Service_View")
    arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "Service_Add")
    arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "Service_ModifyOwn")
    arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "Service_ModifyAll")
    arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "Service_Del")
    arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Complain_View")
    arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "Complain_Add")
    arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "Complain_ModifyOwn")
    arrPurview(13) = CheckPurview_Other(AdminPurview_Others, "Complain_ModifyAll")
    arrPurview(14) = CheckPurview_Other(AdminPurview_Others, "Complain_Del")
    arrPurview(15) = CheckPurview_Other(AdminPurview_Others, "Call_View")
    arrPurview(16) = CheckPurview_Other(AdminPurview_Others, "Call_Add")
    arrPurview(17) = CheckPurview_Other(AdminPurview_Others, "Call_ModifyOwn")
    arrPurview(18) = CheckPurview_Other(AdminPurview_Others, "Call_ModifyAll")
    arrPurview(19) = CheckPurview_Other(AdminPurview_Others, "Dictionary")
    For PurviewIndex = 0 To 4
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Client = True
            Exit For
        End If
    Next
    For PurviewIndex = 5 To 9
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Service = True
            Exit For
        End If
    Next
    For PurviewIndex = 10 To 14
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Complain = True
            Exit For
        End If
    Next
    For PurviewIndex = 15 To 18
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Call = True
            Exit For
        End If
    Next

    If AdminPurview = 1 Or PurviewPassed = True Or arrPurview(19) = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_02.gif' id=menuTitle204 onclick=""new Element.toggle('submenu204')"" style='cursor:hand;'><a href='Admin_Help_CRM.asp' target='main'><span class=glow>�ͻ���ϵ����</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu204'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        If AdminPurview = 1 Or PurviewPassed_Client = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Client.asp target=main>�ͻ�����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Contacter.asp target=main>��ϵ�˹���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or PurviewPassed_Service = True Or PurviewPassed_Call = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Service.asp target=main>�������</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or PurviewPassed_Complain = True Or PurviewPassed_Call = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Complain.asp' target='main'>Ͷ�߹���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(19) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Dictionary.asp target=main>�����ֵ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If FoundInArr(AllModules, "Collection", ",") Then
    If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Collection") = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_05.gif' id=menuTitle210 onclick=""new Element.toggle('submenu210')"" style='cursor:hand;'><a href='Admin_Help_Collection.asp' target='main'><span class=glow>�ɼ�����</span></a></td></tr><tr><td style='display:none' align='right' id='submenu210'><div class=sec_menu style='width:165'><table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Collection.asp?Action=Main target=main>���²ɼ�</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_CollectionManage.asp?Action=ItemManage target=main>��Ŀ����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Filter.asp?Action=main target=main>���˹���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_CollectionHistory.asp?Action=main target=main>�ɼ���ʷ��¼</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_CollectionManage.asp?Action=Import target=main>��Ŀ����</a> | <a href=Admin_CollectionManage.asp?Action=Export target=main>��Ŀ����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Timing.asp?Action=main target=main>��ʱ����</a> | <a href=Admin_Timing.asp?Action=DoMainTiming target='_blank'>������ʱ</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_AreaCollection.asp?Action=Main target=main>����ɼ�����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_9.gif' id=menuTitle301 onclick=""new Element.toggle('submenu301')"" style='cursor:hand;'><a href='Admin_Help_Create.asp' target='main'><span class=glow>��վ���ɹ���</span></a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td style='display:none' align='right' id='submenu301'><div class=sec_menu style='width:165'>" & vbCrLf
Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
If AdminPurview = 1 And FileExt_SiteIndex < 4 Then
    Response.Write "          <tr height=20>" & vbCrLf
    Response.Write "            <td><a href='Admin_CreateSiteIndex.asp' target='main'>������վ��ҳ</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
End If
If AdminPurview = 1 Then
    Response.Write "          <tr height=20>" & vbCrLf
    Response.Write "            <td><a href='Admin_CreateOther.asp' target='main'>������վ�ۺ�����</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
End If
If AdminPurview = 1 And FileExt_SiteSpecial < 4 Then
    Response.Write "          <tr height=20>" & vbCrLf
    Response.Write "            <td><a href='Admin_CreateHTML.asp?Action=SiteSpecial' target='main'>ȫվר�����ɹ���</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
End If
Response.Write strActionLink & vbCrLf
If FileExt_SiteIndex < 4 Or FileExt_SiteSpecial < 4 Or strActionLink <> "" Then
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Timing.asp?Action=main target=main>��ʱ����</a> | <a href=Admin_Timing.asp?Action=DoMainTiming target='_blank'>������ʱ</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
End If
Response.Write "        </table>" & vbCrLf
Response.Write "      </div>" & vbCrLf
Response.Write "        <div  style='width:167'>" & vbCrLf
Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
Response.Write "            <tr>" & vbCrLf
Response.Write "              <td height=5></td>" & vbCrLf
Response.Write "            </tr>" & vbCrLf
Response.Write "          </table>" & vbCrLf
Response.Write "      </div></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Dim PurviewPassed_User
PurviewPassed = False
PurviewPassed_User = False

arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "User_View")
arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "User_ModifyInfo")
arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "User_MofidyPurview")
arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "User_Lock")
arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "User_Del")
arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "User_Update")
arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "User_Money")
arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "User_Point")
arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "User_Valid")
arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "UserGroup")
arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Card")
arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "ConsumeLog")
arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "RechargeLog")
arrPurview(13) = CheckPurview_Other(AdminPurview_Others, "Message")
arrPurview(14) = CheckPurview_Other(AdminPurview_Others, "MailList")
arrPurview(15) = CheckPurview_Other(AdminPurview_Others, "Payment")
arrPurview(16) = CheckPurview_Other(AdminPurview_Others, "Bankroll")
arrPurview(17) = CheckPurview_Other(AdminPurview_Others, "Bank")
For PurviewIndex = 0 To 8
    If arrPurview(PurviewIndex) = True Then
        PurviewPassed = True
        PurviewPassed_User = True
        Exit For
    End If
Next
For PurviewIndex = 9 To 16
    If arrPurview(PurviewIndex) = True Then
        PurviewPassed = True
        Exit For
    End If
Next
If AdminPurview = 1 Or PurviewPassed = True Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/admin_left_11.gif' id=menuTitle203 onclick=""new Element.toggle('submenu203')"" style='cursor:hand;'><a href='Admin_Help_User.asp' target='main'><span class=glow>�û�����</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu203'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    If AdminPurview = 1 Or PurviewPassed_User = True = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_User.asp target=main>��Ա����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(9) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_UserGroup.asp' target='main'>��Ա�����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(10) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Card.asp' target='main'>��ֵ������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(11) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_ConsumeLog.asp' target='main'>��Ա" & PointName & "��ϸ</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(12) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_RechargeLog.asp' target='main'>��Ա��Ч����ϸ</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(13) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Message.asp' target='main'>����Ϣ����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(14) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Maillist.asp' target='main'>�ʼ��б����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(14) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Mail.asp' target='main'>�ʼ����Ĺ���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If

    If AdminPurview = 1 Or arrPurview(15) = True Then
        Response.Write "          <tr height=20>" & vbCrLf
        Response.Write "            <td><a href='Admin_Payment.asp' target='main'>����֧����¼����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(16) = True Then
        Response.Write "          <tr height=20>" & vbCrLf
        Response.Write "            <td><a href='Admin_Bankroll.asp' target='main'>�ʽ���ϸ��ѯ</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or PurviewPassed_User = True = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SpaceManage.asp' target='main'>�ۺϿռ����</a> | <a href='Admin_SpaceManage.asp?Action=Check' target='main'>���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SpaceManage.asp?Action=Kind' target='main'>�ռ�������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Admin.asp target=main>����Ա����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Template.asp?TemplateType=8&TempType=1 target=main>��Աģ��ҳ����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div  style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If
If FoundInArr(AllModules, "SMS", ",") Then
PurviewPassed = False
    arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "SendSMSToMember")
    arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "SendSMSToContacter")
    arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "SendSMSToConsignee")
    arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "SendSMSToOther")
    arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "ViewMessageLog")
    arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "SMS_MessageReceive")
    For PurviewIndex = 0 To 5
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            Exit For
        End If
    Next
    If AdminPurview = 1 Or PurviewPassed = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/admin_left_11.gif' id=menuTitle501 onclick=""new Element.toggle('submenu501')"" style='cursor:hand;'><a href='Admin_Help_SMS.asp' target='main'><span class=glow>�ֻ����Ź���</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu501'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        If AdminPurview = 1 Or arrPurview(0) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SMS.asp?SendTo=Member' target='main'>���͸���Ա</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(1) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SMS.asp?SendTo=Contacter' target='main'>���͸���ϵ��</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(2) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SMS.asp?SendTo=Consignee' target='main'>���͸������е��ջ���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(3) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SMS.asp?SendTo=Other' target='main'>���͸�������</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(4) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SMSLog.asp' target='main'>�鿴���ŷ��ͽ��</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(5) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SMSReceive.asp' target='main'>�鿴���յ��Ķ���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='http://sms.powereasy.net/Member/Recharge.aspx' target='main'>����ͨ��ֵ</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If ShowAdmin_Guest = True Then
    'PurviewPassed = CheckPurview_Other(AdminPurview_Others, "GuestBook")
    If AdminPurview = 1 Or AdminPurview_GuestBook <= 3 Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_4.gif' id=menuTitle202 onclick=""new Element.toggle('submenu202')"" style='cursor:hand;'><a href='Admin_Help_Guest.asp' target='main'><span class=glow>���԰����</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu202'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Passed=False' target=main>��վ�������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Passed=All' target=main>��վ���Թ���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        If AdminPurview = 1 Or AdminPurview_GuestBook < 3 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Action=GKind' target=main>����������</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Channel") = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Channel.asp?Action=Modify&iChannelID=4' target=main>���԰�����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Template") = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=4' target=main>����ģ��ҳ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or AdminPurview_GuestBook < 2 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Action=CreateCode' target=main>��ҳǶ���������</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_CreateXml.asp?Action=GreatNav&ChannelID=4' target=main>������ĿXML����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If


If (SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "All") And FoundInArr(AllModules, "Classroom", ",") Then
    If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Equipment") = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/admin_left_13.gif' id=menuTitle209 onclick=""new Element.toggle('submenu209')"" style='cursor:hand;'><a href='Admin_Help_Classroom.asp' target='main'><span>�ҳ��Ǽǹ���</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu209'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_ClassroomSort.asp' target='main'>�ҳ�����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Equipment.asp' target='main'>�豸����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_ManageRecord.asp?Action=AdminEnrol' target='main'>�̶��ſ�</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        
        Response.Write "            <td height=20><a href='Admin_ManageRecord.asp?Status=All' target='main'>ʹ�õǼǹ���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SearchHistory.asp' target='main'>ʹ�ü�¼��ѯ</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If (SystemEdition = "EPS" Or SystemEdition = "All") And FoundInArr(AllModules, "Sdms", ",") Then
    PurviewPassed = False
    arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "InfoManage")
    arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "ScoreManage")
    arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "TestManage")
    For PurviewIndex = 0 To 2
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            Exit For
        End If
    Next
    If AdminPurview = 1 Or PurviewPassed = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_12.gif' id=menuTitle402 onclick=""new Element.toggle('submenu402')"" style='cursor:hand;'><a href='Admin_Help_Manage.asp' target='main'><span>ѧ��ѧ������</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu402'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        If AdminPurview = 1 Or arrPurview(0) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_StudentInfoManage.asp' target='main'>ѧ����Ϣ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(1) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_ScoreManage.asp' target='main'>ѧ���ɼ�����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(2) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_TestManage.asp' target='main'>���Թ���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SdmsDatabaseManage.asp' target='main'>ѧ�����ݿ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If SystemEdition <> "CMS" And SystemEdition <> "eShop" And FoundInArr(AllModules, "Survey", ",") Then
    PurviewPassed = False
    arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "ViewSurvey")
    arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "AddSurvey")
    arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "ManageSurvey")
    arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "ShowSurveyCountData")
    arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "ManageSurveyTemplate")
    arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "ImportSurveyQuestion")
    arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "ExportSurveyQuestion")
    arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "ViewListQuestion")
    For PurviewIndex = 0 To 7
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            Exit For
        End If
    Next
    If AdminPurview = 1 Or PurviewPassed = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/admin_left_13.gif' id=menuTitle209 onclick=""new Element.toggle('submenu219')"" style='cursor:hand;'><span>�ʾ�������</span></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu219'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        If AdminPurview = 1 Or arrPurview(1) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Survey.asp?Action=AddSurvey' target='main'>�����ʾ�</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(0) = True Or arrPurview(2) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Survey.asp' target='main'>�ʾ����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(2) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Survey.asp?Action=SurveyCode' target='main'>���ô���</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(7) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Survey.asp?Action=ListQuestion' target='main'>��Ŀ�б�</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(2) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Survey.asp?Action=ManageTemplate' target='main'>�ʾ�ģ�����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(5) = True Or arrPurview(6) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Survey.asp?Action=ExportQuestion' target='main'>�ʾ���Ŀ����</a> | <a href='Admin_Survey.asp?Action=ImportQuestion' target='main'>����</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

PurviewPassed = False
arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "Channel")
arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "AD")
arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "FriendSite")
arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "Announce")
arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "Vote")
arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "Counter")
arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "Skin")
arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "Label")
arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "KeyLink")
arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "Rtext")
arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Template")
arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "Bank")
arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "ShowPage")
For PurviewIndex = 0 To 12
    If arrPurview(PurviewIndex) = True Then
        PurviewPassed = True
        Exit For
    End If
Next
If AdminPurview = 1 Or PurviewPassed = True Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_01.gif' id=menuTitle201 onclick=""new Element.toggle('submenu201')"" style='cursor:hand;'><a href='Admin_Help_SiteConfig.asp' target='main'><span class=glow>ϵͳ����</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu201'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_SiteConfig.asp target=main>��վ��Ϣ����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(0) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Channel.asp target=main>��վƵ������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Special.asp target=main>ȫվר�����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(1) = True) And FoundInArr(AllModules, "Advertisement", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Advertisement.asp target=main>��վ������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(2) = True) And FoundInArr(AllModules, "FriendSite", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_FriendSite.asp target=main>�������ӹ���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(3) = True) And FoundInArr(AllModules, "Announce", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Announce.asp target=main>��վ�������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(4) = True) And FoundInArr(AllModules, "Vote", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Vote.asp target=main>��վ�������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(5) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Counter.asp target=main>��վͳ�Ʒ���</a> | <a href=Admin_Counter.asp?Action=ShowConfig target=main>����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(6) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Skin.asp target=main>��վ������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(10) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=0' target='main'>��վͨ��ģ��ҳ����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_TemplateProject.asp?ChannelID=0' target='main'>��վģ�巽������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SourceManage.asp?ChannelID=0&TypeSelect=Keyword' target='main'>��վͨ�ùؼ��ֹ���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(7) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Label.asp' target='main'>�Զ����ǩ����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If	
    If AdminPurview = 1 Or arrPurview(12) = True Then		
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Page.asp' target='main'>�Զ���ҳ�����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Log.asp' target='main'>��վ��־����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SiteCount.asp' target='main'>����Ա������ͳ��</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(8) = True) And FoundInArr(AllModules, "KeyLink", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SourceManage.asp?TypeSelect=KeyLink' target='main'>վ�����ӹ���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(9) = True) And FoundInArr(AllModules, "Rtext", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SourceManage.asp?TypeSelect=Rtext' target='main'>�ַ����˹���</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(11) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Bank.asp' target='main'>�����ʻ�����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_PayPlatform.asp' target='main'>����֧��ƽ̨����</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_City.asp' target='main'>�����������</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_CompareFilesOnline.asp' target='main'>���߱Ƚ���վ�ļ�</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If


If AdminPurview = 1 Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_03.gif' id=menuTitle206 onclick=""new Element.toggle('submenu206')"" style='cursor:hand;'><a href='Admin_Help_Database.asp' target='main'><span class=glow>���ݿ����</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu206'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Backup target=main>�������ݿ�</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Restore target=main>�ָ����ݿ�</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Compact target=main>ѹ�����ݿ�</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Init target=main>ϵͳ��ʼ��</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=SpaceSize target=main>ϵͳ�ռ�ռ��</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div  style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If

Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_04.gif' id=menuTitle208><span>ϵͳ��Ϣ</span> </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align='right'><div class=sec_menu style='width:165'>" & vbCrLf
Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
Response.Write "          <tr>" & vbCrLf
Response.Write "            <td height=20><br>" & vbCrLf
Response.Write "              ��Ȩ���У�&nbsp;<a href='http://www.powereasy.net/' target=_blank>��������</a><br>" & vbCrLf
Response.Write "              ���������&nbsp;<a href='http://www.powereasy.net/' target=_blank>��������</a><BR>" & vbCrLf
Response.Write "              ����֧�֣�&nbsp;<a href='http://bbs.powereasy.net/' target=_blank>������̳</a><br>" & vbCrLf
Response.Write "              <br>" & vbCrLf
Response.Write "            </td>" & vbCrLf
Response.Write "          </tr>" & vbCrLf
Response.Write "        </table>" & vbCrLf
Response.Write "    </div></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf

rsGetAdmin.Close
Set rsGetAdmin = Nothing
Call CloseConn
%>
