<!--#include file="Start.asp"-->
<%
Server.ScriptTimeOut = 9999999

Dim i

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>����SiteWeaver6.8 ���ݿ���������</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link href=""Admin/Admin_STYLE.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body>" & vbCrLf

Action = Trim(request("Action"))
Select Case Action
Case "Upgrade"
    Call Upgrade
Case "Del"
    Call Del
Case Else
    Call Main
End Select
Call CloseConn

Response.Write "</body></html>"

Sub Main()
    Response.Write "<table width=""700""  border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""border"">" & vbCrLf
    Response.Write "  <form name=""myform"" method=""post"" action=""Upgrade.asp"">" & vbCrLf
    Response.Write "  <tr align=""center"" class=""topbg"">" & vbCrLf
    Response.Write "    <td height=""25""><strong>����SiteWeaver6.8 ���ݿ���������</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=""60"" align=right>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""60"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-bottom: 1px solid #999999;"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <strong>�ʺϰ汾��</strong><br>" & vbCrLf
    Response.Write "            &nbsp;&nbsp;���������������ڹٷ������汾����2006ϵ�а汾 ������SiteWeaver6.5 ��SiteWeaver6.6 ��SiteWeaver6.7ϵ�а汾������SiteWeaver6.8�汾�� <br>" & vbCrLf
    Response.Write "            <strong>�������裺</strong><br>" & vbCrLf
    Response.Write "            &nbsp;&nbsp;����ǰ��һ��Ҫ������ϸ���Ķ�����Ĳ������輰ע���������" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "          <td align=""right"" width=""200"" Height=""80"" background=""http://www.powereasy.net/images/logo.gif"">&nbsp;</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg"">" & vbCrLf
    Response.Write "    <td valign=""top "">" & vbCrLf
    Response.Write "      <table width=""90%"" align=""center"" height=""250"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td valign=""top "">" & vbCrLf
    Response.Write "            ȷ�����˽���������ݺ󣬵���[��һ��]������" & vbCrLf
    Response.Write "            <textarea style='width:680px;height:200px' style=""font-size: 9pt;"" readonly>"
    Response.Write " ���������裺" & vbCrLf
    Response.Write " 1�������ļ���Upgrade.asp������ϵͳ��Ŀ¼�¡�" & vbCrLf
    Response.Write " 2��������������뱾�ļ��ĵ�ַ����http://localhost/Upgrade.asp�����б�����" & vbCrLf
    Response.Write " 3�������Ķ���˵����㡰��һ��������ʼ����������" & vbCrLf
    Response.Write " ��ע�����" & vbCrLf
    Response.Write " 1������������ֻ�����ڹٷ������汾�����ݿ��������������������޸İ�������������������" & vbCrLf
    Response.Write " 2��������ֱ���ڷ���������������������ɹ���ɺ�һ��Ҫɾ�����ļ������������ȫ������"
    Response.Write "</textarea>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "      <hr>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td align=""center"">" & vbCrLf
    Response.Write "            <input type=""hidden"" name=""Action"" value=""Upgrade"">" & vbCrLf
    Response.Write "            <input name=""Submit"" type=""submit"" value="" ��һ�� "">" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  </form>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Upgrade()
    'On Error Resume Next

    '����Ƿ�װSMS���ݱ�
    If IsExists("SMSUserName", "PE_Config") = False Then
        '����SMS�ֶ�
        If SystemDatabaseType = "SQL" Then
            CONN.execute ("alter table [PE_Config]  add  SMSUserName [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  SMSKey [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  Mobiles [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  SendMessageToAdminWhenOrder [bit]  Default (0) NOT NULL ")
            CONN.execute ("alter table [PE_Config]  add  SendMessageToMemberWhenPaySuccess [bit]  Default (0) NOT NULL ")
            CONN.execute ("alter table [PE_Config]  add  MessageOfOrder  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfAddRemit  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfAddIncome   [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfAddPayment  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfExchangePoint  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfAddPoint  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfMinusPoint  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfExchangeValid  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfAddValid  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  MessageOfMinusValid  [ntext] COLLATE Chinese_PRC_CI_AS NULL")
        Else
            CONN.execute ("alter table [PE_Config]  add  COLUMN SMSUserName   text(50)")
            CONN.execute ("alter table [PE_Config]  add  COLUMN SMSKey   text(50)")
            CONN.execute ("alter table [PE_Config]  add  COLUMN Mobiles text(255)")
            CONN.execute ("alter table [PE_Config]  add  COLUMN SendMessageToAdminWhenOrder bit")
            CONN.execute ("alter table [PE_Config]  add  COLUMN SendMessageToMemberWhenPaySuccess bit")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfOrder   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfAddRemit   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfAddIncome   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfAddPayment   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfExchangePoint   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfAddPoint   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfMinusPoint   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfExchangeValid   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfAddValid   text")
            CONN.execute ("alter table [PE_Config]  add  COLUMN MessageOfMinusValid   text")
        End If
        Dim rsConfig
        Set rsConfig = Server.CreateObject("ADODB.Recordset")
        rsConfig.open "select * from PE_Config", CONN, 1, 3
        If rsConfig.bof And rsConfig.EOF Then
            rsConfig.AddNew
        End If
        rsConfig("Modules") = rsConfig("Modules") & ",SMS"
        
        rsConfig("SMSUserName") = ""
        rsConfig("SMSKey") = ""
        rsConfig("Mobiles") = ""
        rsConfig("SendMessageToAdminWhenOrder") = True
        rsConfig("SendMessageToMemberWhenPaySuccess") = True
        rsConfig("MessageOfOrder") = "��Ա {$UserName} �� {$InputTime} ����һ���������������Ϊ��{$MoneyTotal}Ԫ��"

        rsConfig("MessageOfOrderConfirm") = "{$ContacterName}���ã����ύ�Ķ�����ȷ�ϡ��밴�����еĽ�����ϵ���ǡ��յ��������ǽ��������ŷ�����{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfReceiptMoney") = "{$ContacterName}���ã��Ѿ��յ��������л��������ڰ��ŷ�����{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfRefund") = "{$ContacterName}���ã������Ѷ����Ķ����������˿����������Ա�ʽ���ϸ�в鿴��ؼ�¼��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfInvoice") = "{$ContacterName}���ã����Ķ����Ѿ����߷�Ʊ��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfDeliver") = "{$ContacterName}���ã����Ķ����Ѿ����������������������ʱ���ʾ���ȡ����������û���յ��뼰ʱ��������ϵ��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfSendCard") = "{$ContacterName}���ã�������ĳ�ֵ����Ϣ���£�{$CardInfo}��{$SiteName}������ظ��˶��ţ�"

        rsConfig("MessageOfAddRemit") = "{$UserName}���ã����㵽{$BankName}��{$Money}Ԫ������յ�������ӵ�����ʻ��С�{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfAddIncome") = "{$UserName} ���ã��Ѿ��������ʻ��������{$Money}Ԫ�������ڵ��ʽ����Ϊ��{$Balance}��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfAddPayment") = "{$UserName} ���ã��Ѵ������ʻ��п۳���{$Money}Ԫ������{$Reason}��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfExchangePoint") = "{$UserName} ���ã��Ѵ������ʻ��п۳���{$Money}Ԫ�����ڶһ�{$Point}��ȯ���ֿ��õ���Ϊ��{$UserPoint}��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfAddPoint") = "{$UserName} ���ã��Ѿ��������ʻ��������{$Point}��ȯ���ֿ��õ���Ϊ��{$UserPoint}��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfMinusPoint") = "{$UserName} ���ã��Ѵ������ʻ��п۳���{$Point}��ȯ������{$Reason}���ֿ��õ���Ϊ��{$UserPoint}��{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfExchangeValid") = "{$UserName} ���ã��Ѵ������ʻ��п۳���{$Money}Ԫ�����ڶһ���Ч��{$Valid}������Ч��ʣ������Ϊ:{$ValidDays}�졣{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfAddValid") = "{$UserName} ���ã��Ѿ��������ʻ����������Ч��{$Valid}������Ч��ʣ������Ϊ:{$ValidDays}�졣{$SiteName}������ظ��˶��ţ�"
        rsConfig("MessageOfMinusValid") = "{$UserName} ���ã��Ѵ����ʻ��п۳���Ч��{$Valid}������{$Reason}������Чʣ������Ϊ:{$ValidDays}�졣{$SiteName}������ظ��˶��ţ�"

        rsConfig.Update
        rsConfig.Close
        Set rsConfig = Nothing
    End If

    Dim haveSurveyTable
    If IsExists("SurveyID", "PE_Survey") = True Then
        haveSurveyTable = True
    Else
        haveSurveyTable = False
    End If

    
    If IsExists("UserName", "PE_ShoppingCarts") = False Then
        '���/����sp6�ֶ�
        If SystemDatabaseType = "SQL" Then
            CONN.execute ("alter table [PE_ComplainItem]  add  Defendant [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_ShoppingCarts]  add  UserName [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  Thumb_BackgroundColor [nvarchar] (10) COLLATE Chinese_PRC_CI_AS NULL")
            CONN.execute ("alter table [PE_Config]  add  PhotoQuality int")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN Meta_Keywords   ntext")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN Meta_Description   ntext")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN Modules   ntext")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN RegFields_MustFill   ntext")
            If haveSurveyTable = True Then
                Dim constraintName
                constraintName = CONN.execute("select b.name as constraintName from syscolumns a,sysobjects b where a.id=object_id('PE_Survey') and b.id=a.cdefault and a.name='IPRepeat' and b.name like 'DF%'")(0)
                CONN.execute ("alter table [PE_Survey] drop constraint " & constraintName)
                CONN.execute ("alter table [PE_Survey] ALTER COLUMN IPRepeat Integer")
            End If
        Else
            CONN.execute ("alter table [PE_ComplainItem]  add  COLUMN Defendant   text(50)")
            CONN.execute ("alter table [PE_ShoppingCarts]  add  COLUMN UserName   text(20)")
            CONN.execute ("alter table [PE_Config]  add  COLUMN Thumb_BackgroundColor text(10)")
            CONN.execute ("alter table [PE_Config]  add  COLUMN PhotoQuality Integer")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN Meta_Keywords   memo")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN Meta_Description   memo")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN Modules   memo")
            CONN.execute ("alter table [PE_Config]  ALTER   COLUMN RegFields_MustFill   memo")
            If haveSurveyTable = True Then
                CONN.execute ("alter table [PE_Survey]  ALTER   COLUMN IPRepeat   Integer")
            End If
        End If
    End If


    '�ʽ���ϸ
    Dim rs, OrderFormNum, trs
    Dim PrefixCoun
    PrefixCoun = Len(Prefix_OrderFormNum)
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select * from PE_BankrollItem where OrderFormID<=0 and Remark<>'' order by ItemID", CONN, 1, 3
    Do While Not rs.EOF
        OrderFormNum = rs("Remark")
        If InStr(OrderFormNum, Prefix_OrderFormNum & "200") > 0 Then
            OrderFormNum = Mid(OrderFormNum, InStr(OrderFormNum, Prefix_OrderFormNum & "200"), 16 + PrefixCoun)
            'response.write OrderFormNum & "<br>"
            If IsNumeric(Right(OrderFormNum, 16)) Then
                Set trs = CONN.execute("select OrderFormID from PE_OrderForm where OrderFormNum='" & OrderFormNum & "'")
                If Not (trs.bof And trs.EOF) Then
                    rs("OrderFormID") = trs(0)
                    rs.Update
                    Response.Write "."
                End If
                Set trs = Nothing
            End If
        End If
        rs.movenext
    Loop
    rs.Close
    Set rs = Nothing
    CONN.execute ("Update PE_Config SET Thumb_BackgroundColor = '#CCCCCC', PhotoQuality = 90")
    CONN.execute ("Update PE_ShoppingCarts SET UserName = ''")
    CONN.execute ("Update PE_City SET AreaCode = '029' WHERE AreaCode = '0910'")
    CONN.execute ("update PE_BankrollItem set [Money]=0-[Money] where Income_Payout=2 And [Money]>0")
    CONN.execute ("update PE_BankrollItem set OrderFormId=0-ItemID where OrderFormID=0")

    If haveSurveyTable = False Then
        Call CreateSurvey
    End If
    
    Dim surveyTemplateCount
    surveyTemplateCount = CONN.execute("Select Count(*) from PE_Template Where ChannelID = 996")(0)
    If surveyTemplateCount <= 0 Then
        Call AddSurveyTemplate
    End If

	Call AddPayPlatformTable

    Call Patch0410

    Dim rsChannel, rsMail, sqlMail, rsCheck, i

    If SystemDatabaseType = "SQL" Then
        If IsExists("fieldlist", "PE_Label") = True Then
            Conn.execute ("alter table [PE_Label]  alter column  fieldlist  [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL")
        End If		
        If IsExists("arrChannelID", "PE_Contacter") = False Then
             Conn.execute ("alter table [PE_Contacter]  add  arrChannelID [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL")
        End If
        If IsExists("ChannelID", "PE_MailChannel") = False Then
            Conn.execute ("create table [PE_MailChannel]( [ChannelID] integer NOT NULL  PRIMARY KEY,[UserID] nvarchar(255),[arrClass] nvarchar(255),[SendNum] integer,[IsUse] bit)")
        End If        	
		For i = 9 to 20
            If IsExists("Select"& i, "PE_Vote") = False Then
                Conn.execute ("alter table [PE_Vote]  add   Select"& i &" nvarchar(255)")
            End If
        Next			
		For i = 9 to 20
            If IsExists("Answer"& i, "PE_Vote") = False Then
                Conn.execute ("alter table [PE_Vote]  add   Answer"& i &" integer")
            End If			
        Next	
		
        If IsExists("VoteNum", "PE_Vote") = False Then
            Conn.execute ("alter table [PE_Vote]  add VoteNum integer")
        End If	
        Conn.execute ("alter table [PE_User]  alter column  [LoginTimes] int")	
        Conn.execute ("alter table [PE_SurveyAnswer]  alter column  AnswerContent [nvarchar] (255)")
        If IsExists("ShowUserModel", "PE_Config") = False Then		
            Conn.execute ("alter table [PE_Config]  add  ShowUserModel bit")	
		End If
        If IsExists("ShowAnonymous", "PE_Config") = False Then				
            Conn.execute ("alter table [PE_Config]  add  ShowAnonymous bit")	
        End If	
        If IsExists("Support", "PE_Comment") = False Then					
            Conn.execute ("alter table [PE_Comment]  add  Support integer")	
        End If
		If IsExists("Opposed", "PE_Comment") = False Then			
            Conn.execute ("alter table [PE_Comment]  add  Opposed integer")		
        End If		
		If IsExists("EnableComment", "PE_Channel") = False Then			
            Conn.execute ("alter table [PE_Channel]  add  EnableComment bit")		
        End If		
		If IsExists("CheckComment", "PE_Channel") = False Then			
            Conn.execute ("alter table [PE_Channel]  add  CheckComment bit")		
        End If							
    Else
        If IsExists("fieldlist", "PE_Label") = True Then
           Conn.execute ("alter table [PE_Label]  alter  fieldlist  text(255)")
        End If
        If IsExists("arrChannelID", "PE_Contacter") = False Then
            Conn.execute ("alter table [PE_Contacter]  add  COLUMN arrChannelID  text(255)")
        End If
        If IsExists("ChannelID", "PE_MailChannel") = False Then
            Conn.execute ("create table [PE_MailChannel]( [ChannelID] integer not null  PRIMARY KEY,[UserID] text(255) null,[arrClass] text(255),[SendNum] integer, [IsUse] bit )")
        End If
		For i = 9 to 20
            If IsExists("Select"& i, "PE_Vote") = False Then
                Conn.execute ("alter table [PE_Vote]  add  COLUMN [Select"& i &"] text(255) null")
            End If
        Next			
		For i = 9 to 20
            If IsExists("Answer"& i, "PE_Vote") = False Then
                Conn.execute ("alter table [PE_Vote]  add  Answer"& i &" integer  null")
            End If			
        Next	
        If IsExists("VoteNum", "PE_Vote") = False Then
            Conn.execute ("alter table [PE_Vote]  add VoteNum integer not null")
        End If	
        Conn.execute ("alter table [PE_User]  alter column  [LoginTimes] int")	
        Conn.execute ("alter table [PE_SurveyAnswer]  alter  AnswerContent  text(255)")
        If IsExists("ShowUserModel", "PE_Config") = False Then		
            Conn.execute ("alter table [PE_Config]  add  ShowUserModel bit")	
		End If
        If IsExists("ShowAnonymous", "PE_Config") = False Then				
            Conn.execute ("alter table [PE_Config]  add  ShowAnonymous bit")	
        End If	
        If IsExists("Support", "PE_Comment") = False Then					
            Conn.execute ("alter table [PE_Comment]  add  Support int")	
        End If
		If IsExists("Opposed", "PE_Comment") = False Then			
            Conn.execute ("alter table [PE_Comment]  add  Opposed int")		
        End If		
		If IsExists("EnableComment", "PE_Channel") = False Then			
            Conn.execute ("alter table [PE_Channel]  add  EnableComment bit")		
        End If		
		If IsExists("CheckComment", "PE_Channel") = False Then			
            Conn.execute ("alter table [PE_Channel]  add  CheckComment bit")		
        End If										
    End If
	
    Dim rsUserGroup, sqlUserGroup
    sqlUserGroup = "Select * from PE_UserGroup Where GroupID = -1"
    Set rsUserGroup = Server.CreateObject("adodb.recordset")	
    rsUserGroup.open sqlUserGroup, CONN, 1, 3
    If rsUserGroup.Bof And rsUserGroup.Eof Then
		Conn.Execute("Insert Into PE_UserGroup (GroupID,GroupName,GroupIntro,GroupType,arrClass_Browse,arrClass_View,arrClass_Input,GroupSetting) Values (-1,'����Ͷ��','����Ͷ���û�Ȩ������',5,'Articlenone,Softnone,Photonone','Articlenone,Softnone,Photonone','Articlenone,Softnone,Photonone','0,0,0,10,1,0,0,1,500,0,1024,100,0,0,0,0,0,100,0,0,0,0,0,0,0,0,0,10,0,0,0')")						
    End If

    				
    Set rsChannel = CONN.execute("select * from PE_Channel where ModuleType = 1")
    sqlMail = "select * from PE_MailChannel"
    Set rsMail = Server.CreateObject("adodb.recordset")
    rsMail.open sqlMail, CONN, 1, 3
    Do While Not rsChannel.EOF
        Set rsCheck = CONN.execute("select * from PE_MailChannel where ChannelID = " & rsChannel("ChannelID"))
        If rsCheck.BOF And rsCheck.EOF Then
            rsMail.addnew
            rsMail("ChannelID") = rsChannel("ChannelID")
            rsMail("UserID") = ""
            rsMail("SendNum") = 10
            rsMail("arrClass") = ""
            rsMail("IsUse") = PE_False
            rsMail.Update
        End If
        rsChannel.movenext
    Loop
    rsMail.Close
    Set rsMail = Nothing
    rsChannel.Close
    Set rsChannel = Nothing
	Call WriteResultInfo
        
End Sub

Sub WriteResultInfo()
    Response.Write "<table id=""Success_Table"" width=""700"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""border"">" & vbCrLf
    Response.Write "  <form name=""myform"" method=""post"" action=""Upgrade.asp"">" & vbCrLf
    Response.Write "  <tr align=""center"" class=""topbg"">" & vbCrLf
    Response.Write "    <td height=""25""><strong>����SiteWeaver6.8 ���ݿ���������</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=""60"" align=right>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""60"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-bottom: 1px solid #999999;"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <strong>������ɣ�</strong><br>" & vbCrLf
    Response.Write "            &nbsp;&nbsp;�ɹ��������ֶ�������SiteWeaver6.8�棡</font><br>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "          <td align=""right"" width=""180"" background=""http://www.powereasy.net/images/logo.gif"">&nbsp;</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg"">" & vbCrLf
    Response.Write "    <td>" & vbCrLf
    Response.Write "      <table width=""90%"" align=""center"" height=""350"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td valign=""top"">" & vbCrLf
    Response.Write "            <br>" & vbCrLf
    Response.Write "            ��ϲ�����ɹ������������ֶ������� SiteWeaver6.8�棡����<br>" & vbCrLf
    Response.Write "            ����ʱ��<span id=""Info_Timer""></span>�롣<br>" & vbCrLf
    Response.Write "            <font color=red>������ֱ���ڷ�����������������������ɾ�����ļ������������ȫ������</font><br><br>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "      <hr>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td align=""center"">" & vbCrLf
    Response.Write "            <input type=""hidden"" name=""Action"" value=""SelectDatabase"">" & vbCrLf
    Response.Write "            <input name=""delfile"" type=""button"" value="" ɾ���˳��� "" onclick=""location='Upgrade.asp?Action=Del'"">" & vbCrLf
    Response.Write "            <input name=""close"" type=""button"" value="" �رմ˴��� "" onclick=""javascript:onclick=window.close()"">" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  </form>" & vbCrLf
    Response.Write "</table>" & vbCrLf
	
End Sub

Sub Del()
    On Error Resume Next
    If fso.FileExists(Server.mappath("Upgrade.asp")) Then
        fso.DeleteFile Server.mappath("Upgrade.asp")
    End If
    If Err.Number <> 0 Then
        Response.Write "<li>ɾ����������Upgrade.asp��ʧ�ܣ�����ԭ��" & Err.Description & "<br>��ʹ��FTPɾ�����ļ���"
        Err.Clear
        Exit Sub
    Else
        Response.Write "<li>ɾ����������Upgrade.asp���ɹ���</li>"
    End If
End Sub

'*********************************************************
'* ���ƣ�IsExists
'* ���ܣ��Ƿ�װSMS���ݱ�(�ֶ���)
'* �÷���IsExists(�ֶ���)
'*********************************************************
Function IsExists(fieldName, tableName)
    On Error Resume Next
    IsExists = True
    CONN.execute ("select " & fieldName & " from " & tableName)

    If Err Then
        IsExists = False
    End If
	Err.Clear
End Function


Sub AddSurveyTemplate()
    If SystemDatabaseType = "SQL" Then
         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,N'�ʾ����Ĭ��ģ��',99,'<html>'+CHAR(13)+CHAR(10)+'<head>'+CHAR(13)+CHAR(10)+N'<title>{$SurveyName}--�ʾ����</title>'+CHAR(13)+CHAR(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+CHAR(13)+CHAR(10)+'{$SurveyJS}'+CHAR(13)+CHAR(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+CHAR(13)+CHAR(10)+'</head>'+CHAR(13)+CHAR(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+CHAR(13)+CHAR(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+CHAR(13)+CHAR(10)+' <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style_01.jpg"" width=""600"" height=""92"" alt=""""></td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td style=""padding:8px;"" background=""../Survey/images/style_03.jpg"">{$GetSurveyForm}</td>'+CHAR(13)+CHAR(10)+' </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style_04.jpg"" width=""600"" height=""17"" alt=""""></td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'</table>'+CHAR(13)+CHAR(10)+'</body>'+CHAR(13)+CHAR(10)+'</html>',1,N'����2006��������',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,N'�ʾ����ģ���',99,'<html>'+CHAR(13)+CHAR(10)+'<head>'+CHAR(13)+CHAR(10)+N'<title>{$SurveyName}--�ʾ����</title>'+CHAR(13)+CHAR(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+CHAR(13)+CHAR(10)+'{$SurveyJS}'+CHAR(13)+CHAR(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+CHAR(13)+CHAR(10)+'<style type=""text/css"">'+CHAR(13)+CHAR(10)+'    body {background-color: #EFEEE2;}'+CHAR(13)+CHAR(10)+'</style>'+CHAR(13)+CHAR(10)+'</head>'+CHAR(13)+CHAR(10)+'<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+CHAR(13)+CHAR(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" id=""__01"">'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style2_01.jpg"" width=""600"" height=""97"" alt=""""></td>'+CHAR(13)+CHAR(10)+'    </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'      <td style=""padding:8px;"" background=""images/style2_02.jpg"">{$GetSurveyForm}</td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'    <tr>'+CHAR(13)+CHAR(10)+'       <td background=""../Survey/images/style2_02.jpg"">&nbsp;</td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style2_03.jpg"" width=""600"" height=""17"" alt=""""></td>'+CHAR(13)+CHAR(10)+' </tr>'+CHAR(13)+CHAR(10)+'</table>'+CHAR(13)+CHAR(10)+'</body>'+CHAR(13)+CHAR(10)+'</html>',0,N'����2006��������',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,N'�ʾ����ģ����',99,'<html>'+CHAR(13)+CHAR(10)+'<head>'+CHAR(13)+CHAR(10)+N'<title>{$SurveyName}--�ʾ����</title>'+CHAR(13)+CHAR(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+CHAR(13)+CHAR(10)+'{$SurveyJS}'+CHAR(13)+CHAR(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+CHAR(13)+CHAR(10)+'</head>'+CHAR(13)+CHAR(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+CHAR(13)+CHAR(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+CHAR(13)+CHAR(10)+'   <tr>'+CHAR(13)+CHAR(10)+'       <td><img src=""../Survey/images/top.jpg"" width=""600"" height=""95""></td>'+CHAR(13)+CHAR(10)+'    </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td  style=""padding:8px; border-left:#999999 solid 1px;border-right:#999999 solid 1px;""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">'+CHAR(13)+CHAR(10)+'          <tr>'+CHAR(13)+CHAR(10)+'            <td >{$GetSurveyForm}</td>'+CHAR(13)+CHAR(10)+'   </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td bgcolor=""#0149a8"">&nbsp;</td>'+CHAR(13)+CHAR(10)+'    </tr>'+CHAR(13)+CHAR(10)+'</table>'+CHAR(13)+CHAR(10)+'</body>'+CHAR(13)+CHAR(10)+'</html>',0,N'����2006��������',1,0)")
    Else
         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,'�ʾ����Ĭ��ģ��',99,'<html>'+Chr(13)+Chr(10)+'<head>'+Chr(13)+Chr(10)+'<title>{$SurveyName}--�ʾ����</title>'+Chr(13)+Chr(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+Chr(13)+Chr(10)+'{$SurveyJS}'+Chr(13)+Chr(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+Chr(13)+Chr(10)+'</head>'+Chr(13)+Chr(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+Chr(13)+Chr(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+Chr(13)+Chr(10)+' <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style_01.jpg"" width=""600"" height=""92"" alt=""""></td>'+Chr(13)+Chr(10)+'    </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td style=""padding:8px;"" background=""../Survey/images/style_03.jpg"">{$GetSurveyForm}</td>'+Chr(13)+Chr(10)+'   </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style_04.jpg"" width=""600"" height=""17"" alt=""""></td>'+Chr(13)+Chr(10)+'    </tr>'+Chr(13)+Chr(10)+'</table>'+Chr(13)+Chr(10)+'</body>'+Chr(13)+Chr(10)+'</html>',1,'����2006��������',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,'�ʾ����ģ���',99,'<html>'+Chr(13)+Chr(10)+'<head>'+Chr(13)+Chr(10)+'<title>{$SurveyName}--�ʾ����</title>'+Chr(13)+Chr(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+Chr(13)+Chr(10)+'{$SurveyJS}'+Chr(13)+Chr(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+Chr(13)+Chr(10)+'<style type=""text/css"">'+Chr(13)+Chr(10)+'    body {background-color: #EFEEE2;}'+Chr(13)+Chr(10)+'</style>'+Chr(13)+Chr(10)+'</head>'+Chr(13)+Chr(10)+'<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+Chr(13)+Chr(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" id=""__01"">'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style2_01.jpg"" width=""600"" height=""97"" alt=""""></td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'      <td style=""padding:8px;"" background=""images/style2_02.jpg"">{$GetSurveyForm}</td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'  <tr>'+Chr(13)+Chr(10)+'     <td background=""../Survey/images/style2_02.jpg"">&nbsp;</td>'+Chr(13)+Chr(10)+'    </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style2_03.jpg"" width=""600"" height=""17"" alt=""""></td>'+Chr(13)+Chr(10)+'   </tr>'+Chr(13)+Chr(10)+'</table>'+Chr(13)+Chr(10)+'</body>'+Chr(13)+Chr(10)+'</html>',0,'����2006��������',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,'�ʾ����ģ����',99,'<html>'+Chr(13)+Chr(10)+'<head>'+Chr(13)+Chr(10)+'<title>{$SurveyName}--�ʾ����</title>'+Chr(13)+Chr(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+Chr(13)+Chr(10)+'{$SurveyJS}'+Chr(13)+Chr(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+Chr(13)+Chr(10)+'</head>'+Chr(13)+Chr(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+Chr(13)+Chr(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+Chr(13)+Chr(10)+'   <tr>'+Chr(13)+Chr(10)+'     <td><img src=""../Survey/images/top.jpg"" width=""600"" height=""95""></td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td  style=""padding:8px; border-left:#999999 solid 1px;border-right:#999999 solid 1px;""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">'+Chr(13)+Chr(10)+'          <tr>'+Chr(13)+Chr(10)+'            <td >{$GetSurveyForm}</td>'+Chr(13)+Chr(10)+' </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td bgcolor=""#0149a8"">&nbsp;</td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'</table>'+Chr(13)+Chr(10)+'</body>'+Chr(13)+Chr(10)+'</html>',0,'����2006��������',1,0)")
    End If
End Sub

Sub CreateSurvey()
    If SystemDatabaseType = "SQL" Then
        CONN.execute ("CREATE TABLE [dbo].[PE_Survey] (" & _
            "[SurveyID] integer IDENTITY (1,1) not null," & _
           "[SurveyName] nvarchar(50) null ," & _
           "[Description] ntext null ," & _
           "[FileName] nvarchar(50) null ," & _
           "[IPRepeat] integer Default (0) null ," & _
           "[CreateDate] datetime Default (getdate()) null ," & _
           "[EndTime] datetime null ," & _
           "[IsOpen] integer Default (0) null ," & _
           "[NeedLogin] integer Default (0) null ," & _
           "[PresentPoint] integer Default (0) null ," & _
           "[LockIPType] integer Default (0) null ," & _
           "[SetIPLock] ntext null ," & _
           "[LockUrl] ntext null ," & _
           "[SetPassword] nvarchar(50) null ," & _
           "[TemplateID] integer Default (0) null " & _
          ") ON [Primary]")

        '[PE_SurveyAnswer]:

        CONN.execute ("CREATE TABLE [dbo].[PE_SurveyAnswer] (" & _
            "[AnswerID] integer IDENTITY (1,1) not null," & _
           "[QuestionID] integer Default (0) null ," & _
           "[AnswerContent] nvarchar(50) null ," & _
           "[VoteAmount] integer Default (0) null ," & _
           "[OrderID] integer Default (0) null " & _
          ") ON [Primary]")

        '[PE_SurveyInput]:

        CONN.execute ("CREATE TABLE [dbo].[PE_SurveyInput] (" & _
           "[QuestionID] integer Default (0) not null ," & _
           "[InputContent] ntext null ," & _
           "[SurveyID] integer Default (0) null " & _
          ") ON [Primary]")

        '[PE_SurveyQuestion]:

        CONN.execute ("CREATE TABLE [dbo].[PE_SurveyQuestion] (" & _
            "[QuestionID] integer IDENTITY (1,1) not null," & _
           "[SurveyID] integer Default (0) null ," & _
           "[QuestionContent] nvarchar(255) null ," & _
           "[QuestionType] tinyint Default (1) null ," & _
           "[InputType] tinyint Default (0) null ," & _
           "[NotEmpty] integer Default (0) null ," & _
           "[DataRight] integer Default (0) null ," & _
           "[OrderID] integer Default (0) null ," & _
           "[ContentLength] integer Default (0) null " & _
          ") ON [Primary]")

        '[PE_Survey]:

        CONN.execute (" Alter TABLE [dbo].[PE_Survey] WITH NOCHECK ADD CONSTRAINT [PK_PE_Survey] Primary Key Clustered ([SurveyID] )  ON [Primary] ")
        CONN.execute ("CREATE INDEX [i_id] on [dbo].[PE_Survey]([SurveyID] ) ON [Primary]")
        CONN.execute ("CREATE INDEX [TemplateID] on [dbo].[PE_Survey]([TemplateID] ) ON [Primary]")

        '[PE_SurveyAnswer]:

        CONN.execute (" Alter TABLE [dbo].[PE_SurveyAnswer] WITH NOCHECK ADD CONSTRAINT [PK_PE_SurveyAnswer] Primary Key Clustered ([AnswerID] )  ON [Primary] ")
        CONN.execute ("CREATE INDEX [item_id] on [dbo].[PE_SurveyAnswer]([QuestionID] ) ON [Primary]")
        CONN.execute ("CREATE INDEX [o_id] on [dbo].[PE_SurveyAnswer]([AnswerID] ) ON [Primary]")
        CONN.execute ("CREATE INDEX [OrderID] on [dbo].[PE_SurveyAnswer]([OrderID] ) ON [Primary]")

        '[PE_SurveyInput]:

        CONN.execute (" Alter TABLE [dbo].[PE_SurveyInput] WITH NOCHECK ADD CONSTRAINT [PK_PE_SurveyInput] Primary Key Clustered ([QuestionID] )  ON [Primary] ")
        CONN.execute ("CREATE INDEX [QuestionID] on [dbo].[PE_SurveyInput]([QuestionID] ) ON [Primary]")
        CONN.execute ("CREATE INDEX [SurveyID] on [dbo].[PE_SurveyInput]([SurveyID] ) ON [Primary]")

        '[PE_SurveyQuestion]:

        CONN.execute (" Alter TABLE [dbo].[PE_SurveyQuestion] WITH NOCHECK ADD CONSTRAINT [PK_PE_SurveyQuestion] Primary Key Clustered ([QuestionID] )  ON [Primary] ")
        CONN.execute ("CREATE INDEX [i_id] on [dbo].[PE_SurveyQuestion]([SurveyID] ) ON [Primary]")
        CONN.execute ("CREATE INDEX [item_id] on [dbo].[PE_SurveyQuestion]([QuestionID] ) ON [Primary]")
        CONN.execute ("CREATE INDEX [OrderID] on [dbo].[PE_SurveyQuestion]([OrderID] ) ON [Primary]")
    Else
        CONN.execute ("CREATE TABLE [PE_Survey] (" & _
            "[SurveyID] integer IDENTITY (1,1) not null," & _
           "[SurveyName] text(50) null ," & _
           "[Description] MEMO null ," & _
           "[FileName] text(50) null ," & _
           "[IPRepeat] integer null Default 0," & _
           "[CreateDate] datetime  null Default now()," & _
           "[EndTime] datetime null ," & _
           "[IsOpen] integer  null Default 0," & _
           "[NeedLogin] integer null Default 0," & _
           "[PresentPoint] integer null Default 0," & _
           "[LockIPType] integer null Default 0," & _
           "[SetIPLock] MEMO null ," & _
           "[LockUrl] MEMO null ," & _
           "[SetPassword] text(50) null ," & _
           "[TemplateID] integer null Default 0" & _
          ") ")

        '[PE_SurveyAnswer]:

        CONN.execute ("CREATE TABLE [PE_SurveyAnswer] (" & _
            "[AnswerID] integer IDENTITY (1,1) not null ," & _
           "[QuestionID] integer null Default 0 ," & _
           "[AnswerContent] text(50) null ," & _
           "[VoteAmount] integer null Default 0 ," & _
           "[OrderID] integer null Default 0 " & _
          ") ")

        '[PE_SurveyInput]:

        CONN.execute ("CREATE TABLE [PE_SurveyInput] (" & _
           "[QuestionID] integer not null Default 0 ," & _
           "[InputContent] MEMO null ," & _
           "[SurveyID] integer null Default 0 " & _
          ") ")

        '[PE_SurveyQuestion]:

        CONN.execute ("CREATE TABLE [PE_SurveyQuestion] (" & _
            "[QuestionID] integer IDENTITY (1,1) not null ," & _
           "[SurveyID] integer null Default 0 ," & _
           "[QuestionContent] text(255) null ," & _
           "[QuestionType] tinyint null Default 1 ," & _
           "[InputType] tinyint null Default 0 ," & _
           "[NotEmpty] integer null Default 0 ," & _
           "[DataRight] integer null Default 0 ," & _
           "[OrderID] integer null Default 0 ," & _
           "[ContentLength] integer null Default 0 " & _
          ") ")

        '[PE_Survey]:

        CONN.execute (" Alter TABLE [PE_Survey] ADD CONSTRAINT [PK_PE_Survey] Primary Key Clustered ([SurveyID] )   ")
        CONN.execute ("CREATE INDEX [i_id] on [PE_Survey]([SurveyID] ) ")
        CONN.execute ("CREATE INDEX [TemplateID] on [PE_Survey]([TemplateID] ) ")

        '[PE_SurveyAnswer]:

        CONN.execute (" Alter TABLE [PE_SurveyAnswer] ADD CONSTRAINT [PK_PE_SurveyAnswer] Primary Key Clustered ([AnswerID] )   ")
        CONN.execute ("CREATE INDEX [item_id] on [PE_SurveyAnswer]([QuestionID] ) ")
        CONN.execute ("CREATE INDEX [o_id] on [PE_SurveyAnswer]([AnswerID] ) ")
        CONN.execute ("CREATE INDEX [OrderID] on [PE_SurveyAnswer]([OrderID] ) ")

        '[PE_SurveyInput]:

        CONN.execute (" Alter TABLE [PE_SurveyInput] ADD CONSTRAINT [PK_PE_SurveyInput] Primary Key Clustered ([QuestionID] )   ")
        CONN.execute ("CREATE INDEX [QuestionID] on [PE_SurveyInput]([QuestionID] ) ")
        CONN.execute ("CREATE INDEX [SurveyID] on [PE_SurveyInput]([SurveyID] ) ")

        '[PE_SurveyQuestion]:

        CONN.execute (" Alter TABLE [PE_SurveyQuestion] ADD CONSTRAINT [PK_PE_SurveyQuestion] Primary Key Clustered ([QuestionID] )   ")
        CONN.execute ("CREATE INDEX [i_id] on [PE_SurveyQuestion]([SurveyID] ) ")
        CONN.execute ("CREATE INDEX [item_id] on [PE_SurveyQuestion]([QuestionID] ) ")
        CONN.execute ("CREATE INDEX [OrderID] on [PE_SurveyQuestion]([OrderID] ) ")
    End If
End Sub

Sub AddPayPlatformTable()
	If IsExists("PlatformID", "PE_PayPlatform") = True Then
		Exit Sub
	End If
    If SystemDatabaseType = "SQL" Then
		CONN.execute ("CREATE TABLE [dbo].[PE_PayPlatform] (" & _
		   "[PlatformID] integer Default (0) not null ," & _
		   "[PlatformName] nvarchar(50) null ," & _
		   "[ShowName] nvarchar(50) null ," & _
		   "[Description] ntext null ," & _
		   "[AccountsID] nvarchar(50) null ," & _
		   "[MD5Key] nvarchar(255) null ," & _
		   "[Rate] float Default (0) null ," & _
		   "[PlusPoundage] bit not null ," & _
		   "[OrderID] integer Default (0) null ," & _
		   "[IsDisabled] bit not null ," & _
		   "[IsDefault] bit not null " & _
		  ") ")
		CONN.execute ("Alter TABLE [dbo].[PE_PayPlatform] WITH NOCHECK ADD CONSTRAINT [PK_PE_PayPlatform] Primary Key Clustered ([PlatformID] )")
		CONN.execute ("CREATE INDEX [IsDisabled] on [dbo].[PE_PayPlatform]([IsDisabled] )")
		CONN.execute ("CREATE INDEX [OrderID] on [dbo].[PE_PayPlatform]([OrderID] )")
	Else
		CONN.execute ("CREATE TABLE [PE_PayPlatform] (" & _
		   "[PlatformID] integer not null Default 0 ," & _
		   "[PlatformName] text(50) null ," & _
		   "[ShowName] text(50) null ," & _
		   "[Description] MEMO null ," & _
		   "[AccountsID] text(50) null ," & _
		   "[MD5Key] text(255) null ," & _
		   "[Rate] float null Default 0 ," & _
		   "[PlusPoundage] bit not null ," & _
		   "[OrderID] integer null Default 0 ," & _
		   "[IsDisabled] bit not null ," & _
		   "[IsDefault] bit not null " & _
		  ") ")
		CONN.execute ("Alter TABLE [PE_PayPlatform] ADD CONSTRAINT [PK_PE_PayPlatform] Primary Key Clustered ([PlatformID] )")
		CONN.execute ("CREATE INDEX [IsDisabled] on [PE_PayPlatform]([IsDisabled] )")
		CONN.execute ("CREATE INDEX [OrderID] on [PE_PayPlatform]([OrderID] )")
	End If

	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (1,'��������','��������','�����֧��ƽ̨�������Ѻã�������','000000','aaaaaaaaaa',1,0,2,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (2,'�й�����֧����','�й�����֧����','���Ǻܺã����Ƽ�','000000','aaaaaaaaaa',1,0,3,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (3,'�Ϻ���Ѹ','�Ϻ���Ѹ','����֧��ƽ̨','000000','aaaaaaaaaa',1,0,4,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (5,'����֧��','����֧��','���Ƽ�','000000','aaaaaaaaaa',1,0,5,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (6,'�׸�ͨ','�׸�ͨ','�����֧��ƽ̨','000000','aaaaaaaaaa',1,0,6,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (7,'����֧��','����֧��','�����֧��ƽ̨','000000','aaaaaaaaaa',1,0,7,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (8,'֧����','֧����','�û����޷�ֱ���õ����ź�����','000000','aaaaaaaaaa',1,0,8,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (9,'��Ǯ֧��','��Ǯ֧��','�����֧��ƽ̨�������Ѻ�','000000','aaaaaaaaaa',1,0,9,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (11,'��Ǯ������','��Ǯ������','','000000','aaaaaaaaaa',1,0,10,0,0)")
	CONN.execute ("Insert into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (12,'֧������ʱ����','֧����','֧������ʱ���ʷ�ʽ','000000','aaaaaaaaaa',1,0,11,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (13,'�Ƹ�ͨ','�Ƹ�ͨ','����ĿǰΨһ�������ѵ�֧��ƽ̨��ǿ���Ƽ���','000000','aaaaaaaaaa',1,0,1,0,0)")
	Conn.Execute("update PE_PayPlatform set IsDisabled=" & PE_True & ",IsDefault=" & PE_False & "")

	Dim rsConfig
	Set rsConfig = Conn.Execute("Select * from PE_Config")
	Conn.Execute("update PE_PayPlatform set IsDisabled=" & PE_False & ",IsDefault=" & PE_True & ",AccountsID='" & rsConfig("PayOnlineShopID") & "',MD5Key='" & rsConfig("PayOnlineKey") & "',Rate=" & rsConfig("PayOnlineRate") & " where PlatformID=" & rsConfig("PayOnlineProvider") & "")
	rsConfig.Close
	Set rsConfig = Nothing

End Sub

Sub Patch0410()
    '�޸��ͻ����ʽ���ϸ��¼���Ӧ���û���һ�µ�����
    Dim rsBankroll,rsUser
    Set rsBankroll = Conn.Execute("select ItemID,UserName from PE_BankrollItem where UserName<>'' and ClientID=0")
    Do While Not rsBankroll.Eof
        Set rsUser = Conn.Execute("select ClientID from PE_User where UserName='" & rsBankroll(1) & "'")
        If Not(rsUser.bof And rsUser.eof) Then
            ClientID = rsUser("ClientID")
            If ClientID > 0 Then
                Conn.Execute("update PE_BankrollItem set ClientID=" & ClientID & " where ItemID=" & rsBankroll(0) & "")
                Response.Write "."
                Response.Flush
            End If
        End If
        rsUser.Close
        Set rsUser=Nothing
        rsBankroll.Movenext
    Loop
    rsBankroll.Close
    Set rsBankroll = Nothing

    '�޸�����Ա��ǰ̨�û������̨�û�����һ��ʱ��ͳ�����ݲ���ȷ������
    Dim rsAdmin
    Set rsAdmin = Conn.Execute("select AdminName,UserName from PE_Admin")
    Do While Not rsAdmin.eof
        If rsAdmin("AdminName") <> rsAdmin("UserName") Then
            Conn.Execute("update PE_Article set Inputer='" & rsAdmin("UserName") & "' where Inputer='" & rsAdmin("AdminName") & "'")
        End If
        rsAdmin.MoveNext
    Loop
    rsAdmin.Close
    Set rsAdmin = Nothing
End Sub
%>

