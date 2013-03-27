<!--#include file="Start.asp"-->
<%
Server.ScriptTimeOut = 9999999

Dim i

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>动易SiteWeaver6.8 数据库升级程序</title>" & vbCrLf
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
    Response.Write "    <td height=""25""><strong>动易SiteWeaver6.8 数据库升级程序</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=""60"" align=right>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""60"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-bottom: 1px solid #999999;"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <strong>适合版本：</strong><br>" & vbCrLf
    Response.Write "            &nbsp;&nbsp;本升级程序适用于官方发布版本动易2006系列版本 ，动易SiteWeaver6.5 ，SiteWeaver6.6 ，SiteWeaver6.7系列版本升级到SiteWeaver6.8版本。 <br>" & vbCrLf
    Response.Write "            <strong>操作步骤：</strong><br>" & vbCrLf
    Response.Write "            &nbsp;&nbsp;升级前请一定要认真仔细的阅读下面的操作步骤及注意事项！！！" & vbCrLf
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
    Response.Write "            确定已了解下面的内容后，单击[下一步]继续。" & vbCrLf
    Response.Write "            <textarea style='width:680px;height:200px' style=""font-size: 9pt;"" readonly>"
    Response.Write " ●升级步骤：" & vbCrLf
    Response.Write " 1、将本文件（Upgrade.asp）复制系统根目录下。" & vbCrLf
    Response.Write " 2、在浏览器中输入本文件的地址，如http://localhost/Upgrade.asp，运行本程序。" & vbCrLf
    Response.Write " 3、认真阅读本说明后点“下一步”，开始升级操作。" & vbCrLf
    Response.Write " ●注意事项：" & vbCrLf
    Response.Write " 1、本升级程序只适用于官方发布版本的数据库升级，不适用于其他修改版或美化版的升级工作。" & vbCrLf
    Response.Write " 2、若您是直接在服务器进行升级，则操作成功完成后，一定要删除此文件！以免带来安全隐患。"
    Response.Write "</textarea>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "      <hr>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td align=""center"">" & vbCrLf
    Response.Write "            <input type=""hidden"" name=""Action"" value=""Upgrade"">" & vbCrLf
    Response.Write "            <input name=""Submit"" type=""submit"" value="" 下一步 "">" & vbCrLf
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

    '检测是否安装SMS数据表
    If IsExists("SMSUserName", "PE_Config") = False Then
        '建立SMS字段
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
        rsConfig("MessageOfOrder") = "会员 {$UserName} 于 {$InputTime} 下了一个订单，订单金额为：{$MoneyTotal}元。"

        rsConfig("MessageOfOrderConfirm") = "{$ContacterName}您好：您提交的订单已确认。请按订单中的金额汇款并联系我们。收到汇款后我们将立即安排发货。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfReceiptMoney") = "{$ContacterName}您好：已经收到您的银行汇款，我们正在安排发货。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfRefund") = "{$ContacterName}您好：我们已对您的订单进行了退款，您可在您会员资金明细中查看相关记录。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfInvoice") = "{$ContacterName}您好：您的订单已经开具发票。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfDeliver") = "{$ContacterName}您好：您的订单已经发货，请留意包裹单并及时到邮局领取。若两周内没有收到请及时和我们联系。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfSendCard") = "{$ContacterName}您好：您购买的充值卡信息如下：{$CardInfo}。{$SiteName}（请勿回复此短信）"

        rsConfig("MessageOfAddRemit") = "{$UserName}您好：您汇到{$BankName}的{$Money}元汇款已收到并已添加到你的帐户中。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfAddIncome") = "{$UserName} 您好：已经给您的帐户中添加了{$Money}元。您现在的资金余额为：{$Balance}。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfAddPayment") = "{$UserName} 您好：已从您的帐户中扣除了{$Money}元，用于{$Reason}。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfExchangePoint") = "{$UserName} 您好：已从您的帐户中扣除了{$Money}元，用于兑换{$Point}点券。现可用点数为：{$UserPoint}。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfAddPoint") = "{$UserName} 您好：已经给您的帐户中添加了{$Point}点券。现可用点数为：{$UserPoint}。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfMinusPoint") = "{$UserName} 您好：已从您的帐户中扣除了{$Point}点券，用于{$Reason}。现可用点数为：{$UserPoint}。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfExchangeValid") = "{$UserName} 您好：已从您的帐户中扣除了{$Money}元，用于兑换有效期{$Valid}。现有效期剩余天数为:{$ValidDays}天。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfAddValid") = "{$UserName} 您好：已经给您的帐户中添加了有效期{$Valid}。现有效期剩余天数为:{$ValidDays}天。{$SiteName}（请勿回复此短信）"
        rsConfig("MessageOfMinusValid") = "{$UserName} 您好：已从您帐户中扣除有效期{$Valid}，用于{$Reason}。现有效剩余天数为:{$ValidDays}天。{$SiteName}（请勿回复此短信）"

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
        '添加/更改sp6字段
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


    '资金明细
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
		Conn.Execute("Insert Into PE_UserGroup (GroupID,GroupName,GroupIntro,GroupType,arrClass_Browse,arrClass_View,arrClass_Input,GroupSetting) Values (-1,'匿名投稿','匿名投稿用户权限设置',5,'Articlenone,Softnone,Photonone','Articlenone,Softnone,Photonone','Articlenone,Softnone,Photonone','0,0,0,10,1,0,0,1,500,0,1024,100,0,0,0,0,0,100,0,0,0,0,0,0,0,0,0,10,0,0,0')")						
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
    Response.Write "    <td height=""25""><strong>动易SiteWeaver6.8 数据库升级程序</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=""60"" align=right>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""60"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-bottom: 1px solid #999999;"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <strong>升级完成：</strong><br>" & vbCrLf
    Response.Write "            &nbsp;&nbsp;成功将数据字段升级至SiteWeaver6.8版！</font><br>" & vbCrLf
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
    Response.Write "            恭喜您，成功将动易数据字段升级至 SiteWeaver6.8版！！！<br>" & vbCrLf
    Response.Write "            共耗时：<span id=""Info_Timer""></span>秒。<br>" & vbCrLf
    Response.Write "            <font color=red>若您是直接在服务器进行升级，则请立即删除此文件！以免带来安全隐患。</font><br><br>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>"
    Response.Write "      <hr>" & vbCrLf
    Response.Write "      <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td align=""center"">" & vbCrLf
    Response.Write "            <input type=""hidden"" name=""Action"" value=""SelectDatabase"">" & vbCrLf
    Response.Write "            <input name=""delfile"" type=""button"" value="" 删除此程序 "" onclick=""location='Upgrade.asp?Action=Del'"">" & vbCrLf
    Response.Write "            <input name=""close"" type=""button"" value="" 关闭此窗口 "" onclick=""javascript:onclick=window.close()"">" & vbCrLf
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
        Response.Write "<li>删除升级程序（Upgrade.asp）失败，错误原因：" & Err.Description & "<br>请使用FTP删除此文件。"
        Err.Clear
        Exit Sub
    Else
        Response.Write "<li>删除升级程序（Upgrade.asp）成功！</li>"
    End If
End Sub

'*********************************************************
'* 名称：IsExists
'* 功能：是否安装SMS数据表(字段名)
'* 用法：IsExists(字段名)
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
         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,N'问卷调查默认模板',99,'<html>'+CHAR(13)+CHAR(10)+'<head>'+CHAR(13)+CHAR(10)+N'<title>{$SurveyName}--问卷调查</title>'+CHAR(13)+CHAR(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+CHAR(13)+CHAR(10)+'{$SurveyJS}'+CHAR(13)+CHAR(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+CHAR(13)+CHAR(10)+'</head>'+CHAR(13)+CHAR(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+CHAR(13)+CHAR(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+CHAR(13)+CHAR(10)+' <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style_01.jpg"" width=""600"" height=""92"" alt=""""></td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td style=""padding:8px;"" background=""../Survey/images/style_03.jpg"">{$GetSurveyForm}</td>'+CHAR(13)+CHAR(10)+' </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style_04.jpg"" width=""600"" height=""17"" alt=""""></td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'</table>'+CHAR(13)+CHAR(10)+'</body>'+CHAR(13)+CHAR(10)+'</html>',1,N'动易2006海蓝方案',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,N'问卷调查模板二',99,'<html>'+CHAR(13)+CHAR(10)+'<head>'+CHAR(13)+CHAR(10)+N'<title>{$SurveyName}--问卷调查</title>'+CHAR(13)+CHAR(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+CHAR(13)+CHAR(10)+'{$SurveyJS}'+CHAR(13)+CHAR(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+CHAR(13)+CHAR(10)+'<style type=""text/css"">'+CHAR(13)+CHAR(10)+'    body {background-color: #EFEEE2;}'+CHAR(13)+CHAR(10)+'</style>'+CHAR(13)+CHAR(10)+'</head>'+CHAR(13)+CHAR(10)+'<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+CHAR(13)+CHAR(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" id=""__01"">'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style2_01.jpg"" width=""600"" height=""97"" alt=""""></td>'+CHAR(13)+CHAR(10)+'    </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'      <td style=""padding:8px;"" background=""images/style2_02.jpg"">{$GetSurveyForm}</td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'    <tr>'+CHAR(13)+CHAR(10)+'       <td background=""../Survey/images/style2_02.jpg"">&nbsp;</td>'+CHAR(13)+CHAR(10)+'  </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td>'+CHAR(13)+CHAR(10)+'           <img src=""../Survey/images/style2_03.jpg"" width=""600"" height=""17"" alt=""""></td>'+CHAR(13)+CHAR(10)+' </tr>'+CHAR(13)+CHAR(10)+'</table>'+CHAR(13)+CHAR(10)+'</body>'+CHAR(13)+CHAR(10)+'</html>',0,N'动易2006海蓝方案',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,N'问卷调查模板三',99,'<html>'+CHAR(13)+CHAR(10)+'<head>'+CHAR(13)+CHAR(10)+N'<title>{$SurveyName}--问卷调查</title>'+CHAR(13)+CHAR(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+CHAR(13)+CHAR(10)+'{$SurveyJS}'+CHAR(13)+CHAR(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+CHAR(13)+CHAR(10)+'</head>'+CHAR(13)+CHAR(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+CHAR(13)+CHAR(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+CHAR(13)+CHAR(10)+'   <tr>'+CHAR(13)+CHAR(10)+'       <td><img src=""../Survey/images/top.jpg"" width=""600"" height=""95""></td>'+CHAR(13)+CHAR(10)+'    </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td  style=""padding:8px; border-left:#999999 solid 1px;border-right:#999999 solid 1px;""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">'+CHAR(13)+CHAR(10)+'          <tr>'+CHAR(13)+CHAR(10)+'            <td >{$GetSurveyForm}</td>'+CHAR(13)+CHAR(10)+'   </tr>'+CHAR(13)+CHAR(10)+'  <tr>'+CHAR(13)+CHAR(10)+'       <td bgcolor=""#0149a8"">&nbsp;</td>'+CHAR(13)+CHAR(10)+'    </tr>'+CHAR(13)+CHAR(10)+'</table>'+CHAR(13)+CHAR(10)+'</body>'+CHAR(13)+CHAR(10)+'</html>',0,N'动易2006海蓝方案',1,0)")
    Else
         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,'问卷调查默认模板',99,'<html>'+Chr(13)+Chr(10)+'<head>'+Chr(13)+Chr(10)+'<title>{$SurveyName}--问卷调查</title>'+Chr(13)+Chr(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+Chr(13)+Chr(10)+'{$SurveyJS}'+Chr(13)+Chr(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+Chr(13)+Chr(10)+'</head>'+Chr(13)+Chr(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+Chr(13)+Chr(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+Chr(13)+Chr(10)+' <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style_01.jpg"" width=""600"" height=""92"" alt=""""></td>'+Chr(13)+Chr(10)+'    </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td style=""padding:8px;"" background=""../Survey/images/style_03.jpg"">{$GetSurveyForm}</td>'+Chr(13)+Chr(10)+'   </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style_04.jpg"" width=""600"" height=""17"" alt=""""></td>'+Chr(13)+Chr(10)+'    </tr>'+Chr(13)+Chr(10)+'</table>'+Chr(13)+Chr(10)+'</body>'+Chr(13)+Chr(10)+'</html>',1,'动易2006海蓝方案',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,'问卷调查模板二',99,'<html>'+Chr(13)+Chr(10)+'<head>'+Chr(13)+Chr(10)+'<title>{$SurveyName}--问卷调查</title>'+Chr(13)+Chr(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+Chr(13)+Chr(10)+'{$SurveyJS}'+Chr(13)+Chr(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+Chr(13)+Chr(10)+'<style type=""text/css"">'+Chr(13)+Chr(10)+'    body {background-color: #EFEEE2;}'+Chr(13)+Chr(10)+'</style>'+Chr(13)+Chr(10)+'</head>'+Chr(13)+Chr(10)+'<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+Chr(13)+Chr(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" id=""__01"">'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style2_01.jpg"" width=""600"" height=""97"" alt=""""></td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'      <td style=""padding:8px;"" background=""images/style2_02.jpg"">{$GetSurveyForm}</td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'  <tr>'+Chr(13)+Chr(10)+'     <td background=""../Survey/images/style2_02.jpg"">&nbsp;</td>'+Chr(13)+Chr(10)+'    </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td>'+Chr(13)+Chr(10)+'         <img src=""../Survey/images/style2_03.jpg"" width=""600"" height=""17"" alt=""""></td>'+Chr(13)+Chr(10)+'   </tr>'+Chr(13)+Chr(10)+'</table>'+Chr(13)+Chr(10)+'</body>'+Chr(13)+Chr(10)+'</html>',0,'动易2006海蓝方案',1,0)")

         Conn.Execute("Insert Into PE_Template (ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,IsDefaultInProject,Deleted) Values (996,'问卷调查模板三',99,'<html>'+Chr(13)+Chr(10)+'<head>'+Chr(13)+Chr(10)+'<title>{$SurveyName}--问卷调查</title>'+Chr(13)+Chr(10)+'<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">'+Chr(13)+Chr(10)+'{$SurveyJS}'+Chr(13)+Chr(10)+'<link href=""css.css"" rel=""stylesheet"" type=""text/css"">'+Chr(13)+Chr(10)+'</head>'+Chr(13)+Chr(10)+'<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">'+Chr(13)+Chr(10)+'<table width=""600"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">'+Chr(13)+Chr(10)+'   <tr>'+Chr(13)+Chr(10)+'     <td><img src=""../Survey/images/top.jpg"" width=""600"" height=""95""></td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td  style=""padding:8px; border-left:#999999 solid 1px;border-right:#999999 solid 1px;""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">'+Chr(13)+Chr(10)+'          <tr>'+Chr(13)+Chr(10)+'            <td >{$GetSurveyForm}</td>'+Chr(13)+Chr(10)+' </tr>'+Chr(13)+Chr(10)+'    <tr>'+Chr(13)+Chr(10)+'     <td bgcolor=""#0149a8"">&nbsp;</td>'+Chr(13)+Chr(10)+'  </tr>'+Chr(13)+Chr(10)+'</table>'+Chr(13)+Chr(10)+'</body>'+Chr(13)+Chr(10)+'</html>',0,'动易2006海蓝方案',1,0)")
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

	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (1,'网银在线','网银在线','不错的支付平台，界面友好，不掉单','000000','aaaaaaaaaa',1,0,2,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (2,'中国在线支付网','中国在线支付网','不是很好，不推荐','000000','aaaaaaaaaa',1,0,3,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (3,'上海环迅','上海环迅','老牌支付平台','000000','aaaaaaaaaa',1,0,4,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (5,'西部支付','西部支付','不推荐','000000','aaaaaaaaaa',1,0,5,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (6,'易付通','易付通','不错的支付平台','000000','aaaaaaaaaa',1,0,6,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (7,'云网支付','云网支付','不错的支付平台','000000','aaaaaaaaaa',1,0,7,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (8,'支付宝','支付宝','用户会无法直接拿到卡号和密码','000000','aaaaaaaaaa',1,0,8,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (9,'快钱支付','快钱支付','不错的支付平台，界面友好','000000','aaaaaaaaaa',1,0,9,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (11,'快钱神州行','快钱神州行','','000000','aaaaaaaaaa',1,0,10,0,0)")
	CONN.execute ("Insert into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (12,'支付宝即时到帐','支付宝','支付宝即时到帐方式','000000','aaaaaaaaaa',1,0,11,0,0)")
	CONN.execute ("Insert Into PE_PayPlatform (PlatformID,PlatformName,ShowName,Description,AccountsID,MD5Key,Rate,PlusPoundage,OrderID,IsDisabled,IsDefault) Values (13,'财付通','财付通','国内目前唯一免手续费的支付平台！强力推荐！','000000','aaaaaaaaaa',1,0,1,0,0)")
	Conn.Execute("update PE_PayPlatform set IsDisabled=" & PE_True & ",IsDefault=" & PE_False & "")

	Dim rsConfig
	Set rsConfig = Conn.Execute("Select * from PE_Config")
	Conn.Execute("update PE_PayPlatform set IsDisabled=" & PE_False & ",IsDefault=" & PE_True & ",AccountsID='" & rsConfig("PayOnlineShopID") & "',MD5Key='" & rsConfig("PayOnlineKey") & "',Rate=" & rsConfig("PayOnlineRate") & " where PlatformID=" & rsConfig("PayOnlineProvider") & "")
	rsConfig.Close
	Set rsConfig = Nothing

End Sub

Sub Patch0410()
    '修复客户的资金明细记录与对应的用户不一致的问题
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

    '修复管理员的前台用户名与后台用户名不一致时，统计数据不正确的问题
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

