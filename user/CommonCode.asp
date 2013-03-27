<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If CheckUserLogined() = False Then
	Call CloseConn
    Response.Redirect "User_Login.asp"
End If

Call GetUser(UserName)

ChannelID=PE_Clng(Trim(Request("ChannelID")))

Sub GetClass()
	ClassName = ""
	RootID = 0
	ParentID = 0
	Depth = 0
	ParentPath = "0"
	Child = 0
	arrChildID = ""
    If ClassID > 0 Then
        Dim tClass
        Set tClass = Conn.Execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ��</li>"
        Else
            ClassName = tClass(0)
            RootID = tClass(1)
            ParentID = tClass(2)
            Depth = tClass(3)
            ParentPath = tClass(4)
            Child = tClass(5)
            arrChildID = tClass(6)
        End If
        Set tClass = Nothing
    End If
End Sub

'**************************************************
'��������CheckUser_ChannelInput
'��  �ã�����û��Ƿ��д�Ƶ��Ȩ��(�����û���̨��������ж�)
'��  ����iChannelID ----Ƶ��ID
'        ChannelDir ---- Ƶ��Ŀ¼
'        arrClassInput ----��Ŀ����Ȩ��
'����ֵ��True ---- ��Ȩ��
'**************************************************
Function CheckUser_ChannelInput()
    Dim rs
    CheckUser_ChannelInput = False
    If FoundInArr(arrClass_Input, ChannelDir & "all", ",") = True Then
        CheckUser_ChannelInput = True
    Else
        Set rs = Conn.Execute("select ClassID from PE_Class where ChannelID=" & ChannelID)
        Do While Not rs.EOF
            If InStr("," & arrClass_Input & ",", "," & rs("ClassID") & ",") > 0 Then
                CheckUser_ChannelInput = True
                Exit Do
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
End Function

'**************************************************
'��������User_GetClass_Option
'��  �ã���ʾ�û���Ŀ�����˵�
'��  ����ShowType ----��ʾ����
'        CurrentID ----��ǰ��ĿID
'����ֵ���û���Ŀ�����˵�
'**************************************************
Function User_GetClass_Option(ShowType, CurrentID)
    Dim rsClass, sqlClass, strClass_Option, tmpDepth, i, ClassNum
    Dim arrShowLine(20)
    Dim CheckParentPath, PurviewChecked
    
    ClassNum = 1
    CurrentID = PE_CLng(CurrentID)
    
    sqlClass = "Select * from PE_Class where ChannelID=" & ChannelID & " And ClassType=1 order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strClass_Option = strClass_Option & "<option value=''>���������Ŀ</option>"
    Else
        Do While Not rsClass.EOF
            ClassNum = ClassNum + 1
            tmpDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            If ShowType = 1 Then
                strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
            Else
                If rsClass("ParentID") > 0 Then
                    CheckParentPath = ChannelDir & "all," & rsClass("ParentPath") & "," & rsClass("ClassID") & "," & rsClass("arrChildID")
                Else
                    CheckParentPath = ChannelDir & "all," & rsClass("ClassID") & "," & rsClass("arrChildID")
                End If

                If CheckPurview_Class(arrClass_Input, CheckParentPath) = True Then
                    PurviewChecked = True
                    If rsClass("Child") > 0 And rsClass("EnableAdd") = False And rsClass("ClassID") <> CurrentID Then
                        strClass_Option = strClass_Option & "<option value='0'"
                    Else
                        strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                    End If
                Else
                    PurviewChecked = False
                End If
            End If
            If ShowType = 1 Or PurviewChecked = True Then
                If CurrentID = 0 Then
                    If ClassNum = 1 Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                Else
                    If rsClass("ClassID") = CurrentID Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                End If
                strClass_Option = strClass_Option & ">"
                
                If tmpDepth > 0 Then
                    For i = 1 To tmpDepth
                        strClass_Option = strClass_Option & "&nbsp;&nbsp;"
                    Next
                End If
                strClass_Option = strClass_Option & rsClass("ClassName")
                strClass_Option = strClass_Option & "</option>"
            
                ClassNum = ClassNum + 1
            End If
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    User_GetClass_Option = strClass_Option
End Function



Function UserMenu()
				Dim strUserMenu
				strUserMenu = strUserMenu & "<script language='JavaScript1.2' type='text/JavaScript'>" & vbCrLf
				strUserMenu = strUserMenu & "stm_bm(['uueoehr',400,'','" & InstallDir & "images/blank.gif',0,'','',0,0,0,0,0,1,0,0]);" & vbCrLf
				strUserMenu = strUserMenu & "stm_bp('p0',[0,4,0,0,2,2,0,0,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,0,0,'#000000','transparent','',3,0,0,'#000000']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_ai('p0i0',[0,'','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i1','p0i0',[0,'��Ա������ҳ','','',-1,-1,0,'Index.asp','_self','Index.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i3','p0i0',[0,'��Ϣ����','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_bp('p1',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				Dim sqlChannel, rsChannel
				sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False
%>
    <!--#include file="../Include/PowerEasy.Edition.asp"-->
<%
				Select Case SystemEdition
				Case "CMS", "eShop"
				    sqlChannel = sqlChannel & " and ModuleType<4"
				Case "GPS", "EPS", "ECS"
				    sqlChannel = sqlChannel & " and (ModuleType<4 or ModuleType=8)"
				Case "IPS"
				    sqlChannel = sqlChannel & " and (ModuleType<4 or ModuleType=6 or ModuleType=7)"
				Case "All"
				    sqlChannel = sqlChannel & " and (ModuleType<4 or ModuleType>5)"
				End Select
				sqlChannel = sqlChannel & " order by OrderID"
				Set rsChannel = Conn.Execute(sqlChannel)
				Do While Not rsChannel.EOF
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
				    Case 6
				        ModuleName = "Supply"
				    Case 7
				        ModuleName = "House"
				    Case 8
				        ModuleName = "Job"
				    End Select
				    If ChannelID = 998 Then
				        strUserMenu = strUserMenu & "stm_aix('p1i0','p0i0',[0,'" & ChannelName & "����','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_bpx('p2','p1',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				        Dim rsHouseClass
						Set rsHouseClass = Conn.Execute("select * from PE_HouseConfig")
				        Do While Not rsHouseClass.EOF
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'����" & rsHouseClass("ClassName") & "��Ϣ','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ClassID=" & rsHouseClass("ClassID") & "&Action=Add','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'����" & rsHouseClass("ClassName") & "��Ϣ','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ClassID=" & rsHouseClass("ClassID") & "&Action=Manage','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				            rsHouseClass.MoveNext
				        Loop
				        strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				    End If
				    If ChannelID = 997 Then
				        strUserMenu = strUserMenu & "stm_aix('p1i0','p0i0',[0,'�ҵļ�������','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_bpx('p2','p1',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'��ѯְλ��Ϣ','','',-1,-1,0,'../Job/Searchresult.asp','_self','../Job/Searchresult.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'ά���ҵļ���','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Resume','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'�������ְλ','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Supply','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Supply' ,'','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				    End If
				    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = False And ChannelID <> 997 And ChannelID <> 998 Then '���Ӳ���ʾ����������
				        strUserMenu = strUserMenu & "stm_aix('p1i0','p0i0',[0,'" & ChannelName & "����','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_bpx('p2','p1',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				        If CheckUser_ChannelInput() = True Then
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'���" & ChannelShortName & "','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        End If
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'����ӵ�" & ChannelShortName & "','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'���ղص�" & ChannelShortName & "','','',-1,-1,0,'User_Favorite.asp?ChannelID=" & ChannelID & "','_self','User_Favorite.asp?ChannelID=" & ChannelID & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'�����۵�" & ChannelShortName & "','','',-1,-1,0,'User_Comment.asp?ChannelID=" & ChannelID & "','_self','User_Comment.asp?ChannelID=" & ChannelID & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        If rsChannel("ModuleType") = 1 Then
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'ǩ�����¹���','','',-1,-1,0,'User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Receive&Passed=All','_self','User_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Receive&Passed=All','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        End If
				        strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				    Else
				    End If
				    rsChannel.MoveNext
				Loop
				rsChannel.Close
				Set rsChannel = Nothing
				If FoundInArr(AllModules, "Classroom", ",") Then
				    strUserMenu = strUserMenu & "stm_aix('p1i0','p0i0',[0,'�ҳ�ʹ�õǼ�','','',-1,-1,0,'User_Enrol.asp','_self','User_Enrol.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				End If
				    
				strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				Dim rsChannel_Shop, NoShow_Shop
				Set rsChannel_Shop = Conn.Execute("select Disabled from PE_Channel where ModuleType=5")
				If Not (rsChannel_Shop.bof And rsChannel_Shop.EOF) Then
				    NoShow_Shop = rsChannel_Shop(0)
				Else
				    NoShow_Shop = True
				End If
				If NoShow_Shop = False Then
				    strUserMenu = strUserMenu & "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p0i4','p0i0',[0,'�̳ǹ���','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				    If PE_Clng(UserSetting(30)) = 1 Then
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p1i0',[0,'������Ʒ','','',-1,-1,0,'User_Wholesale.asp','_self','User_Wholesale.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    End If
				    If GroupType = 4 Then
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�Ҵ���Ķ���','','',-1,-1,0,'User_Order.asp?OrderType=1','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵĶ��˵�','','',-1,-1,0,'User_Bill.asp','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��Ͷ�߼�¼','','',-1,-1,0,'User_Complain.asp','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    End If
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵĶ���','','',-1,-1,0,'User_Order.asp','_self','User_Order.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵĹ��ﳵ','','',-1,-1,0,'../Shop/ShoppingCart.asp','_blank','../Shop/ShoppingCart.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'���ղص���Ʒ','','',-1,-1,0,'User_Favorite.asp?ChannelID=1000','_self','User_Favorite.asp?ChannelID=1000','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�����۵���Ʒ','','',-1,-1,0,'User_Comment.asp?ChannelID=1000','_self','User_Comment.asp?ChannelID=1000','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'����֧��','','',-1,-1,0,'../PayOnline/PayOnline.asp','_blank','../PayOnline/PayOnline.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'����֧����ѯ','','',-1,-1,0,'User_Payment.asp','_self','User_Payment.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ʽ���ϸ��ѯ','','',-1,-1,0,'User_Bankroll.asp','_self','User_Bankroll.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'���ع�������','','',-1,-1,0,'User_Down.asp','_self','User_Down.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��ȡ�����ֵ��','','',-1,-1,0,'User_Exchange.asp?Action=GetCard','_self','User_Exchange.asp?Action=GetCard','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				End If
				strUserMenu = strUserMenu & "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i5','p0i0',[0,'����Ϣ����','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Inbox','_self','User_Message.asp?Action=Manage&ManageType=Inbox','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'׫д����Ϣ','','',-1,-1,0,'User_Message.asp?Action=New','_self','User_Message.asp?Action=New','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ռ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Inbox','_self','User_Message.asp?Action=Manage&ManageType=Inbox','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ݸ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Outbox','_self','User_Message.asp?Action=Manage&ManageType=Outbox','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ѷ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=IsSend','_self','User_Message.asp?Action=Manage&ManageType=IsSend','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ϼ���','','',-1,-1,0,'User_Message.asp?Action=Manage&ManageType=Recycle','_self','User_Message.asp?Action=Manage&ManageType=Recycle','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i7','p0i0',[0,'��ֵ����','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				If UserSetting(18) = 1 Then
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�һ�" & PointName & "','','',-1,-1,0,'User_Exchange.asp?Action=Exchange','_self','User_Exchange.asp?Action=Exchange','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				End If
				If UserSetting(19) = 1 Then
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�һ���Ч��','','',-1,-1,0,'User_Exchange.asp?Action=Valid','_self','User_Exchange.asp?Action=Valid','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				End If
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��ֵ����ֵ','','',-1,-1,0,'User_Exchange.asp?Action=Recharge','_self','User_Exchange.asp?Action=Recharge','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				If UserSetting(20) = 1 Then
				    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'����" & PointName & "','','',-1,-1,0,'User_Exchange.asp?Action=SendPoint','_self','User_Exchange.asp?Action=SendPoint','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				End If
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'����֧����ѯ','','',-1,-1,0,'User_Payment.asp','_self','User_Payment.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ʽ���ϸ��ѯ','','',-1,-1,0,'User_Bankroll.asp','_self','User_Bankroll.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'" & PointName & "��ϸ��ѯ','','',-1,-1,0,'User_ConsumeLog.asp','_self','User_ConsumeLog.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��Ч����ϸ��ѯ','','',-1,-1,0,'User_RechargeLog.asp','_self','User_RechargeLog.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				If UserSetting(25) = 1 Then
				    strUserMenu = strUserMenu & "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_aix('p0i8','p0i0',[0,'�ҵľۺ�','','',-1,-1,0,'User_Space.asp?Action=Manage','_self','User_Space.asp?Action=Manage','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				    Dim rsspace, rsitem
				    Set rsspace = Conn.Execute("select top 1 Passed from PE_Space where Type=1 and UserID=" & UserID)
				    If rsspace.bof And rsspace.EOF Then
				        strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'����ۺϿռ�','','',-1,-1,0,'User_Space.asp?Action=Add','_self','User_Space.asp?Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				    Else
				        If rsspace("Passed") = True Then
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��������Ŀ','','',-1,-1,0,'User_Space.asp?Action=Add','_self','User_Space.asp?Action=Add','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				            Set rsitem = Conn.Execute("select ID,Name,Type from PE_Space where (Type>=3 and Type<=7) and Passed=" & PE_True & " and UserID=" & UserID & " order by Type desc")
				            Do While Not rsitem.EOF
				                Select Case rsitem("Type")
				                Case 3
				                    strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'д��־','','',-1,-1,0,'User_SpaceDiary.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceDiary.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵ���־����','','',-1,-1,0,'User_SpaceDiary.asp?ID=" & rsitem("ID") & "','_self','User_SpaceDiary.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				                Case 4
				                    strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�������','','',-1,-1,0,'User_SpaceMusic.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceMusic.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵ����ֹ���','','',-1,-1,0,'User_SpaceMusic.asp?ID=" & rsitem("ID") & "','_self','User_SpaceMusic.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				                Case 5
				                    strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�������','','',-1,-1,0,'User_SpaceBook.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceBook.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵ�ͼ�����','','',-1,-1,0,'User_SpaceBook.asp?ID=" & rsitem("ID") & "','_self','User_SpaceBook.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				                Case 6
				                    strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�����ͼƬ','','',-1,-1,0,'User_SpacePhoto.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpacePhoto.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵ�ͼƬ����','','',-1,-1,0,'User_SpacePhoto.asp?ID=" & rsitem("ID") & "','_self','User_SpacePhoto.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				                Case 7
				                    strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'" & rsitem("Name") & "','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'���������','','',-1,-1,0,'User_SpaceLink.asp?Action=Add&ID=" & rsitem("ID") & "','_self','User_SpaceLink.asp?Action=Add&ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ҵ����ӹ���','','',-1,-1,0,'User_SpaceLink.asp?ID=" & rsitem("ID") & "','_self','User_SpaceLink.asp?ID=" & rsitem("ID") & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				                    strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				                End Select
				                rsitem.MoveNext
				            Loop
				            Set rsitem = Nothing
				            If UserSetting(28) = 1 Then
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�����ռ��ʽ','','',-1,-1,0,'User_Space.asp?Action=Template','_self','User_Space.asp?Action=Template','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				            End If
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�鿴�ҵľۺ�','','',-1,-1,0,'../Space/" & UserName & UserID & "/','_blank','../Space/" & UserName & UserID & "','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        Else
				            strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�ۺϿռ������...','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				        End If
				    End If
				End If
				strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i6','p0i0',[0,'�û�����','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,4,0,6,2,3,6,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'�����б�','','',-1,-1,0,'','_self','','','','',6,0,0,'" & InstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_bpx('p2','p0',[1,2,-2,-3,2,3,0,7,100,'filter:Glow(Color=#000000, Strength=3)',4,'',23,50,2,4,'#999999','#0089F7','',3,1,1,'#ACA899']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��Ա�б�','','',-1,-1,0,'User_Friend.asp','_self','User_Friend.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��ӳ�Ա','','',-1,-1,0,'User_Friend.asp?Action=AddFriend','_self','User_Friend.asp?Action=AddFriend','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'��������','','',-1,-1,0,'User_Friend.asp?Action=CreateNewGroup','_self','User_Friend.asp?Action=CreateNewGroup','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p2i0','p0i0',[0,'�������','','',-1,-1,0,'User_Friend.asp?Action=ManageGroup','_self','User_Friend.asp?Action=ManageGroup','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'�޸�����','','',-1,-1,0,'User_Info.asp?Action=ModifyPwd','_self','User_Info.asp?Action=ModifyPwd','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'�޸���Ϣ','','',-1,-1,0,'User_Info.asp?Action=Modify','_self','User_Info.asp?Action=Modify','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				If UserType = 0 Then
				    strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'ע���ҵ���ҵ','','',-1,-1,0,'User_Info.asp?Action=RegCompany','_self','User_Info.asp?Action=RegCompany','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				End If
				strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'�ʼ����Ĺ���','','',-1,-1,0,'User_mailreg.asp','_self','User_mailreg.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_aix('p0i0','p0i0',[0,'�˳���¼','','',-1,-1,0,'User_Logout.asp','_self','User_Logout.asp','','','',0,0,0,'','',0,0,0,0,1,'#F1F2EE',1,'#CCCCCC',1,'','',3,3,0,0,'#FFFFF7','#FF0000','#ffffff','#ffff00','9pt ����','9pt ����']);" & vbCrLf
				strUserMenu = strUserMenu & "stm_ep();" & vbCrLf
				strUserMenu = strUserMenu & "stm_em();" & vbCrLf
				strUserMenu = strUserMenu & "</script>" & vbCrLf
				UserMenu = strUserMenu
End Function


Function ShowMessageBox() 
Dim tMessageID, rsMessage, strMessage
If request("Action") <> "ReadInbox" Then
    Set rsMessage = Conn.Execute("select Min(Id) from PE_Message where incept='" & UserName & "'and delR=0 and flag=0 and IsSend=1")
    If IsNull(rsMessage(0)) Then
        tMessageID = 0
    Else
        tMessageID = rsMessage(0)
    End If
    Set rsMessage = Nothing
    If tMessageID > 0 Then
        strMessage = strMessage &"<script LANGUAGE='JavaScript'>" & vbCrLf
        strMessage = strMessage & "var url = 'User_ReadMessage.asp?MessageID=" & tMessageID & "';" & vbCrLf
        strMessage = strMessage & "window.open (url, 'newmessage', 'height=440, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')" & vbCrLf
       strMessage = strMessage & "</script>" & vbCrLf
    End If
End If
ShowMessageBox = strMessage
End Function
%>
