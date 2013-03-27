<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="../Include/PowerEasy.CreateThumb.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 3   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim HtmlDir, PurviewChecked
Dim ManageType, Status, MyStatus, arrStatus
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created
Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview

Dim ArticleID

Dim PayStatus
Dim IncludePic, UploadFiles, DefaultPicUrl, IsThumb
Dim ArticlePro1, ArticlePro2, ArticlePro3, ArticlePro4
Dim VoteID

If ChannelID = 0 Then
    Response.Write "Ƶ��������ʧ��"
    Call CloseConn
    Response.End
End If
If ModuleType <> 1 Then
    Response.Write "<li>ָ����Ƶ��ID���ԣ�</li>"
    Call CloseConn
    Response.End
End If
ModuleName = "Article"
SheetName = "PE_Article"


If AdminPurview = 1 Then
    MyStatus = 3
Else
    Select Case CheckLevel
    Case 0, 1
        MyStatus = 3
    Case 2
        If AdminPurview_Channel <= 2 Then
            MyStatus = 3
        Else
            MyStatus = 2
        End If
    Case 3
        MyStatus = 4 - AdminPurview_Channel
    End Select
End If
arrStatus = Array("�����", "һ��ͨ��", "����ͨ��", "����ͨ��")

HtmlDir = InstallDir & ChannelDir

ManageType = Trim(Request("ManageType"))
Status = Trim(Request("Status"))
Created = Trim(Request("Created"))
OnTop = Trim(Request("OnTop"))
IsElite = Trim(Request("IsElite"))
IsHot = Trim(Request("IsHot"))
ClassID = PE_CLng(Trim(Request("ClassID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
ArticleID = Trim(Request("ArticleID"))
PayStatus = Trim(Request("PayStatus")) '����֧��״̬

If Action = "" Then
    Action = "Manage"
End If
If Status = "" Then
    Status = 9
Else
    Status = PE_CLng(Status) '����״̬   9�����������£�-1�����ݸ壬0��������ˣ�1��������ˣ�-2�����˸�
End If
If IsValidID(ArticleID) = False Then
    ArticleID = ""
End If
If PayStatus = "" Then
    PayStatus = "False"
End If

FileName = "Admin_Article.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
strFileName = FileName & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&Field=" & strField & "&keyword=" & Keyword

If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
ArticlePro1 = XmlText("Article", "ArticlePro1", "[ͼ��]")
ArticlePro2 = XmlText("Article", "ArticlePro2", "[��ͼ]")
ArticlePro3 = XmlText("Article", "ArticlePro3", "[�Ƽ�]")
ArticlePro4 = XmlText("Article", "ArticlePro4", "[ע��]")


If Action = "ExportExcel" Then
    Call ExportExcel
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
    Response.End
End If
Response.Write "<html><head><title>" & ChannelShortName & "����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
If Action = "Add" Or Action = "Modify" Then
    Response.Write "<script src=""../JS/prototype.js""></script>" & vbCrLf
    Response.Write "<script src=""../JS/checklogin.js""></script>" & vbCrLf
End If
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
If Action = "Preview" Then
    Call Preview
    Call CloseConn
    Response.End
End If
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Dim strTitle
strTitle = ChannelName & "����----"
Select Case Action
Case "Add"
    strTitle = strTitle & "���" & ChannelShortName
Case "Modify"
    strTitle = strTitle & "�޸�" & ChannelShortName
Case "Check"
    strTitle = strTitle & "���" & ChannelShortName
Case "SaveAdd", "SaveModify", "SaveModifyAsAdd"
    strTitle = strTitle & "����" & ChannelShortName
Case "Move"
    strTitle = strTitle & "�ƶ�" & ChannelShortName
Case "Preview", "Show"
    strTitle = strTitle & "Ԥ��" & ChannelShortName
Case "Batch", "DoBatch"
    strTitle = strTitle & "�����޸�" & ChannelShortName & "����"
Case "MoveToClass"
    strTitle = strTitle & "�����ƶ�" & ChannelShortName
Case "BatchReplace"
    strTitle = strTitle & "�����滻" & ChannelShortName
Case "AddToSpecial"
    strTitle = strTitle & "���" & ChannelShortName & "��ר��"
Case "MoveToSpecial"
    strTitle = strTitle & "�ƶ�" & ChannelShortName & "��ר��"
Case "Manage"
    Select Case ManageType
    Case "Check"
        strTitle = strTitle & ChannelShortName & "���"
    Case "PayMoney"
        strTitle = strTitle & ChannelShortName & "��ѹ���"
    Case "HTML"
        strTitle = strTitle & ChannelShortName & "����"
    Case "Recyclebin"
        strTitle = strTitle & ChannelShortName & "����վ����"
    Case "Special"
        strTitle = strTitle & "ר��" & ChannelShortName & "����"
    Case Else
        strTitle = strTitle & ChannelShortName & "������ҳ"
    End Select
End Select
Call ShowPageTitle(strTitle, 10111)

Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>��������</strong></td><td colspan='5'>"
Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Status=9'>" & ChannelShortName & "������ҳ</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>���" & ChannelShortName & "</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=Check&Status=0'>���" & ChannelShortName & "</a>"
If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=Special'>ר��" & ChannelShortName & "����</a>"
End If
If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=Recyclebin&Status=9'>" & ChannelShortName & "����վ����</a>"
End If
If FoundInArr(arrEnabledTabs, "Copyfee", ",") = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ManageType=PayMoney&PayStatus=False' target=main>��ѹ���</a>"
End If
If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=HTML&Status=1'><b>����HTML����</b></a>"
End If
Response.Write "</td></tr>" & vbCrLf

If Action = "Manage" Then
    Response.Write "<form name='form3' method='Post' action='" & strFileName & "'><tr class='tdbg'>"
    Response.Write "  <td width='70' height='30' ><strong>" & ChannelShortName & "ѡ�</strong></td><td>"
    If ManageType = "PayMoney" Then
        Response.Write "<input name='PayStatus' type='radio' onclick='submit();' " & RadioValue(PayStatus, "False") & ">δ֧����ѵ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp"
        Response.Write "<input name='PayStatus' type='radio' onclick='submit();' " & RadioValue(PayStatus, "True") & ">��֧����ѵ�" & ChannelShortName
    ElseIf ManageType = "HTML" Then
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "") & ">����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "False") & ">δ���ɵ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp"
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "True") & ">�����ɵ�" & ChannelShortName
    Else
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 9) & ">����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, -1) & ">�ݸ�&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 0) & ">�����&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 1) & ">�����&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, -2) & ">�˸�"
        Response.Write "</td><td>"
        Response.Write "<input name='OnTop' type='checkbox' onclick='submit()' " & RadioValue(OnTop, "True") & "> �̶�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='IsElite' type='checkbox' onclick='submit()' " & RadioValue(IsElite, "True") & "> �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='IsHot' type='checkbox' onclick='submit()' " & RadioValue(IsHot, "True") & "> ����" & ChannelShortName
    End If
    Response.Write "</td></tr></form>" & vbCrLf
End If
Response.Write "</table>" & vbCrLf

strFileName = strFileName & "&Status=" & Status & "&Created=" & Created & "&PayStatus=" & PayStatus & "&OnTop=" & OnTop & "&IsElite=" & IsElite & "&IsHot=" & IsHot

Select Case Action
Case "Add"
    Call Add
Case "Modify", "Check"
    Call Modify
Case "SaveAdd", "SaveModify", "SaveModifyAsAdd"
    Call SaveArticle
Case "SetOnTop", "CancelOnTop", "SetElite", "CancelElite", "SetPassed", "CancelPassed", "Reject"
    Call SetProperty
Case "Show"
    Call Show
Case "Del"
    Call Del
Case "ConfirmDel"
    Call ConfirmDel
Case "ClearRecyclebin"
    Call ClearRecyclebin
Case "Restore"
    Call Restore
Case "RestoreAll"
    Call RestoreAll
Case "DelFile"
    Call DelFile
Case "Batch"
    Call Batch
Case "DoBatch"
    Call DoBatch
Case "MoveToClass"
    Call ShowForm_MoveToClass
Case "MoveToSpecial"
    Call ShowForm_MoveToSpecial
Case "AddToSpecial"
    Call ShowForm_AddToSpecial
Case "DoMoveToClass"
    Call DoMoveToClass
Case "DoMoveToSpecial"
    Call DoMoveToSpecial
Case "DoAddToSpecial"
    Call DoAddToSpecial
Case "DelFromSpecial"
    Call DelFromSpecial
Case "BatchReplace"
    Call BatchReplace
Case "DoBatchReplace"
    Call DoBatchReplace
Case "Manage"
    Call main
Case "ConfirmPay"
    Call ConfirmPay
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    If ManageType = "HTML" And UseCreateHTML = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Ƶ�������˲�����HTML�����Բ��ý������ɹ���</li>"
        Exit Sub
    End If

    Dim arrUser, i, NotReceiveUser
    Dim rsArticleList, sql, Querysql
    PurviewChecked = False
    If ClassID = 0 Then
        If strField = "" And AdminPurview = 2 And AdminPurview_Channel = 3 And ManageType <> "My" Then
            If ManageType = "Check" Or ManageType = "PayMoney" Then
                If arrClass_Check = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�Բ�����û���ڴ�Ƶ�����" & ChannelShortName & "��Ȩ�ޣ�</li>"
                    Exit Sub
                End If
                Set tClass = Conn.Execute("select top 1 ClassID,ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " and ClassID In (" & DelRightComma(arrClass_Check) & ")")
            Else
                If arrClass_Manage = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�Բ�����û���ڴ�Ƶ������" & ChannelShortName & "��Ȩ�ޣ�</li>"
                    Exit Sub
                End If
                Set tClass = Conn.Execute("select top 1 ClassID,ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " and ClassID In (" & DelRightComma(arrClass_Manage) & ")")
            End If
            If tClass.BOF And tClass.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Բ�����û���ڴ�Ƶ���Ĺ���Ȩ�ޣ�</li>"
            Else
                ClassID = tClass(0)
                ClassName = tClass(1)
                RootID = tClass(2)
                ParentID = tClass(3)
                Depth = tClass(4)
                ParentPath = tClass(5)
                Child = tClass(6)
                arrChildID = tClass(7)
                PurviewChecked = True
                ClassPurview = tClass(8)
                ParentDir = tClass(9)
                ClassDir = tClass(10)
            End If
        End If
    ElseIf ClassID = -1 Then
        If AdminPurview = 1 Or (AdminPurview = 2 And AdminPurview_Channel < 3) Then PurviewChecked = True
    ElseIf ClassID > 0 Then
        Set tClass = Conn.Execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ</li>"
        Else
            ClassName = tClass(0)
            RootID = tClass(1)
            ParentID = tClass(2)
            Depth = tClass(3)
            ParentPath = tClass(4)
            Child = tClass(5)
            arrChildID = tClass(6)
            ClassPurview = tClass(7)
            ParentDir = tClass(8)
            ClassDir = tClass(9)
        End If
        Set tClass = Nothing
    End If

    If FoundErr = True Then Exit Sub


    If PurviewChecked = False Then
        If ManageType = "Check" Or ManageType = "PayMoney" Then
            If ParentID > 0 Then
                PurviewChecked = CheckPurview_Class(arrClass_Check, ParentPath & "," & ClassID)
            Else
                PurviewChecked = CheckPurview_Class(arrClass_Check, ClassID)
            End If
        Else
            If ParentID > 0 Then
                PurviewChecked = CheckPurview_Class(arrClass_Manage, ParentPath & "," & ClassID)
            Else
                PurviewChecked = CheckPurview_Class(arrClass_Manage, ClassID)
            End If
        End If
    End If

    Call ShowJS_Manage(ChannelShortName)
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    If ManageType = "Special" Then
        Response.Write "<tr class='title'>"
        Response.Write "  <td height='22'>" & GetSpecialList() & "</td></tr>" & vbCrLf
    Else
        Response.Write "  <tr class='title'>"
        Response.Write "    <td height='22'>" & GetRootClass() & "</td>"
        Response.Write "  </tr>" & GetChild_Root() & ""
    End If
    Response.Write "</table><br>"

    Select Case ManageType
    Case "Check"
        Call ShowContentManagePath(ChannelShortName & "���")
    Case "PayMoney"
        Call ShowContentManagePath("��ѹ���")
    Case "Receive"
        Call ShowContentManagePath(ChannelShortName & "ǩ�չ���")
    Case "HTML"
        Call ShowContentManagePath(ChannelShortName & "����")
    Case "Recyclebin"
        Call ShowContentManagePath(ChannelShortName & "����վ����")
    Case "Special"
        Call ShowContentManagePath("ר��" & ChannelShortName & "����")
    Case Else
        Call ShowContentManagePath(ChannelShortName & "����")
    End Select

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='Admin_Article.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    If ManageType = "Special" Then
        Response.Write "        <td width='120' align='center'><strong>����ר��</strong></td>"
    End If
    Response.Write "            <td align='center' ><strong>" & ChannelShortName & "����</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>¼����</strong></td>"
      '��Ӹ�ѹ������
    If ManageType = "PayMoney" Then
        Response.Write "        <td width='60' align='center'><strong>����</strong></td>"
        Response.Write "        <td width='80' align='center'><strong>���������</strong></td>"
        Response.Write "        <td width='60' align='center'><strong>���</strong></td>"
        Response.Write "        <td width='40' align='center'><strong>��֧��</strong></td>"
        If PayStatus = "True" Then
            Response.Write "        <td width='60' align='center' ><strong>֧������</strong></td>"
        Else
            Response.Write "        <td width='60' align='center' ><strong>¼������</strong></td>"
        End If
    Else
        Response.Write "            <td width='40' align='center' ><strong>�����</strong></td>"
        Response.Write "            <td width='80' align='center' ><strong>" & ChannelShortName & "����</strong></td>"
        Response.Write "            <td width='60' align='center' ><strong>���״̬</strong></td>"
    End If
    If UseCreateHTML > 0 And ObjInstalled_FSO = True And ManageType <> "Special" Then
        Response.Write "            <td width='40' align='center' ><strong>������</strong></td>"
    End If
    If ManageType = "Check" Then
        Response.Write "            <td width='120' align='center' ><strong>��˲���</strong></td>"
    ElseIf ManageType = "PayMoney" Then
        If PayStatus = "False" Then
            Response.Write "            <td width='60' align='center' ><strong>��Ѳ���</strong></td>"
        End If
    ElseIf ManageType = "HTML" Then
        Response.Write "            <td width='180' align='center' ><strong>����HTML����</strong></td>"
    ElseIf ManageType = "Recyclebin" Then
        Response.Write "            <td width='100' align='center' ><strong>����վ����</strong></td>"
    ElseIf ManageType = "Special" Then
        Response.Write "            <td width='100' align='center' ><strong>ר��������</strong></td>"
    Else
        Response.Write "            <td width='150' align='center' ><strong>����������</strong></td>"
    End If
    Response.Write "          </tr>"

    If ManageType = "Special" Then
        sql = "select top " & MaxPerPage & " I.InfoID,I.SpecialID,A.ArticleID,SP.SpecialName,A.Title,A.Keyword,A.Author,A.UpdateTime,A.Inputer,"
        sql = sql & "A.CopyFrom,A.DefaultPicUrl,A.IncludePic,A.PaginationType,A.Receive,A.ReceiveUser,A.Received,"
        sql = sql & "A.Hits,A.OnTop,A.Elite,A.Status,A.Stars,A.InfoPoint,A.VoteID"
        sql = sql & " from PE_Article A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.ArticleID=I.ItemID "
    Else
        If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
            sql = "select top " & MaxPerPage & " A.ClassID,A.ArticleID,A.Title,A.Keyword,A.Author,A.UpdateTime,A.Inputer,"
            sql = sql & "A.CopyFrom,A.IncludePic,A.DefaultPicUrl,A.PaginationType,A.Receive,A.ReceiveUser,A.Received,"
            sql = sql & "A.Hits,A.OnTop,A.Elite,A.Status,A.Stars,A.InfoPoint,A.Beneficiary,A.IsPayed,A.CopyMoney,A.PayDate,A.VoteID"
            sql = sql & " from PE_Article A "
        Else
            sql = "select top " & MaxPerPage & " A.ClassID,A.ArticleID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,A.Title,A.Keyword,A.Author,A.UpdateTime,A.Inputer,"
            sql = sql & "A.CopyFrom,A.IncludePic,A.DefaultPicUrl,A.PaginationType,A.Receive,A.ReceiveUser,A.Received,"
            sql = sql & "A.Hits,A.OnTop,A.Elite,A.Status,A.Stars,A.InfoPoint,A.Beneficiary,A.IsPayed,A.CopyMoney,A.PayDate,A.VoteID"
            sql = sql & " from PE_Article A left join PE_Class C on A.ClassID=C.ClassID "
        End If
    End If
    
    Querysql = " where A.ChannelID=" & ChannelID
    If ManageType = "Special" Then
        Querysql = Querysql & " and I.ModuleType=" & ModuleType
    End If
    If ManageType = "Receive" Then
        Querysql = Querysql & " and A.Receive=" & PE_True & ""
    End If
    If ManageType = "Recyclebin" Then
        Querysql = Querysql & " and A.Deleted=" & PE_True & ""
    Else
        Querysql = Querysql & " and A.Deleted=" & PE_False & ""
    End If
    If ManageType = "HTML" Then
        If Created = "False" Then
            Querysql = Querysql & " and (A.CreateTime<=A.UpdateTime or A.CreateTime is Null)"
        ElseIf Created = "True" Then
            Querysql = Querysql & " and A.CreateTime>A.UpdateTime"
        End If
        Querysql = Querysql & " and A.Status=3" '������Ϊ�����ʱ���������ɹ����г���
    ElseIf ManageType = "PayMoney" Then
       '����Ǹ�ѹ����������е���ͨ����鵫��ȴû�б�ɾ���ͼ����ѵ�����
        If PayStatus = "False" Then
            Querysql = Querysql & " and A.Status=3 and A.CopyMoney>0 and  A.IsPayed=" & PE_False & "" '��ѯ�������������û�б�֧��������
        ElseIf PayStatus = "True" Then
            Querysql = Querysql & " and A.Status=3 and A.CopyMoney>0 and A.IsPayed=" & PE_True & "" '��ѯ���������ˣ�����֧����������
        End If
    Else
        Select Case Status
        Case -2 '�˸�
            Querysql = Querysql & " and A.Status=-2"
        Case -1 '�ݸ�
            Querysql = Querysql & " and A.Status=-1"
        Case 0  '�����
            Querysql = Querysql & " and A.Status>=0 and A.Status<" & MyStatus
        Case 1  '�����
            Querysql = Querysql & " and A.Status>=" & MyStatus
        Case Else
            Querysql = Querysql & " and A.Status>-1"
        End Select
        If OnTop = "True" Then
            Querysql = Querysql & " and A.OnTop=" & PE_True & ""
        End If
        If IsElite = "True" Then
            Querysql = Querysql & " and A.Elite=" & PE_True & ""
        End If
        If IsHot = "True" Then
            Querysql = Querysql & " and A.Hits>=" & HitsOfHot & ""
        End If
    End If

    If ClassID <> 0 Then
        If Child > 0 Then
            Querysql = Querysql & " and A.ClassID in (" & arrChildID & ")"
        Else
            Querysql = Querysql & " and A.ClassID=" & ClassID
        End If
    End If
    If SpecialID > 0 Then
        Querysql = Querysql & " and I.SpecialID=" & SpecialID
    End If
    If ManageType = "My" Then
        Querysql = Querysql & " and A.Inputer='" & UserName & "' "
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "Title"
            Querysql = Querysql & " and A.Title like '%" & Keyword & "%' "
        Case "Content"
            Querysql = Querysql & " and A.Content like '%" & Keyword & "%' "
        Case "Author"
            Querysql = Querysql & " and A.Author like '%" & Keyword & "%' "
        Case "Inputer"
            Querysql = Querysql & " and A.Inputer='" & Keyword & "' "
        Case "Editor"
            Querysql = Querysql & " and A.Editor='" & Keyword & "' "
        Case "UpdateTime"
            Querysql = Querysql & " and DateDiff(" & PE_DatePart_D & ",A.UpdateTime,'" & Keyword & "')=0 "
        Case "Keyword"
            Querysql = Querysql & " and A.Keyword like '%|" & Keyword & "|%' "
        Case "ID"
            Querysql = Querysql & " and A.ArticleID=" & PE_Clng(Keyword) & " "
        Case Else
            Querysql = Querysql & " and A.Title like '%" & Keyword & "%' "
        End Select
    End If
    If ManageType = "Special" Then
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_InfoS I inner join PE_Article A on I.ItemID=A.ArticleID " & Querysql)(0))
    Else
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_Article A " & Querysql)(0))
    End If
    If CurrentPage < 1 Then
        CurrentPage = 1
    End If
    If (CurrentPage - 1) * MaxPerPage > totalPut Then
        If (totalPut Mod MaxPerPage) = 0 Then
            CurrentPage = totalPut \ MaxPerPage
        Else
            CurrentPage = totalPut \ MaxPerPage + 1
        End If
    End If
    If CurrentPage > 1 Then
        If ManageType = "Special" Then
            Querysql = Querysql & " and I.InfoID < (select min(InfoID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " I.InfoID from PE_InfoS I inner join PE_Article A on I.ItemID=A.ArticleID " & Querysql & " order by I.InfoID desc) as QueryArticle)"
        Else
            Querysql = Querysql & " and A.ArticleID < (select min(ArticleID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " A.ArticleID from PE_Article A " & Querysql & " order by A.ArticleID desc) as QueryArticle)"
        End If
    End If
    If ManageType = "Special" Then
        sql = sql & Querysql & " order by I.InfoID desc"
    Else
        sql = sql & Querysql & " order by A.ArticleID desc"
    End If

    Set rsArticleList = Server.CreateObject("ADODB.Recordset")
    rsArticleList.Open sql, Conn, 1, 1
    If rsArticleList.BOF And rsArticleList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>"
        If ClassID > 0 Then
            Response.Write "����Ŀ��������Ŀ��û���κ�"
        Else
            Response.Write "û���κ�"
        End If
        If ManageType = "PayMoney" Then
            Select Case PayStatus
            Case "True"
                Response.Write "<font color=blue>�Ѹ����</font>��" & ChannelShortName & "��"
            Case "False"
                Response.Write "<font color=green>δ֧�����</font>" & ChannelShortName & "��"
            'Case Else
              '  Response.Write "��Ҫ֧����ѵ�" & ChannelShortName & "��"
            End Select
        Else
            Select Case Status
            Case -2
                Response.Write "�˸�"
            Case -1
                Response.Write "�ݸ�"
            Case 0
                Response.Write "<font color=blue>�����</font>��" & ChannelShortName & "��"
            Case 1
                Response.Write "<font color=green>�����</font>��" & ChannelShortName & "��"
            Case Else
                Response.Write ChannelShortName & "��"
            End Select
        End If
        Response.Write "<br><br></td></tr>"
    Else
        Dim ArticleNum, ArticlePath
        ArticleNum = 0
        Do While Not rsArticleList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            If ManageType = "Special" Then
                Response.Write "        <td width='30' align='center'><input name='InfoID' type='checkbox' onclick='CheckItem(this)' id='InfoID' value='" & rsArticleList("InfoID") & "'></td>"
                Response.Write "        <td width='25' align='center'>" & rsArticleList("InfoID") & "</td>"
                Response.Write "        <td width='120' align='center'>"
                If rsArticleList("SpecialID") > 0 Then
                    Response.Write "<a href='" & FileName & "&SpecialID=" & rsArticleList("SpecialID") & "'>" & rsArticleList("SpecialName") & "</a>"
                Else
                    Response.Write "&nbsp;"
                End If
                Response.Write "</td>"
            Else
                Response.Write "        <td width='30' align='center'><input name='ArticleID' type='checkbox' onclick='CheckItem(this)' id='ArticleID' value='" & rsArticleList("ArticleID") & "'></td>"
                Response.Write "        <td width='25' align='center'>" & rsArticleList("ArticleID") & "</td>"
            End If
            Response.Write "        <td>"
            If ManageType <> "Special" Then
                If rsArticleList("ClassID") <> ClassID And ClassID <> -1 Then
                    Response.Write "<a href='" & FileName & "&ClassID=" & rsArticleList("ClassID") & "'>["
                    If rsArticleList("ClassName") <> "" Then
                        Response.Write rsArticleList("ClassName")
                    Else
                        Response.Write "<font color='gray'>�������κ���Ŀ</font>"
                    End If
                    Response.Write "]</a>&nbsp;"
                End If
            End If
            
            Select Case rsArticleList("IncludePic")
                Case 1
                    Response.Write "<font color=blue>" & ArticlePro1 & "</font>"
                Case 2
                    Response.Write "<font color=blue>" & ArticlePro2 & "</font>"
                Case 3
                    Response.Write "<font color=blue>" & ArticlePro3 & "</font>"
                Case 4
                    Response.Write "<font color=blue>" & ArticlePro4 & "</font>"
            End Select
            
            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsArticleList("ArticleID") & "'"
            Response.Write " title='��&nbsp;&nbsp;&nbsp;&nbsp;�⣺" & rsArticleList("Title") & vbCrLf & "��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�" & rsArticleList("Author") & vbCrLf & "ת �� �ԣ�" & rsArticleList("CopyFrom") & vbCrLf & "����ʱ�䣺" & rsArticleList("UpdateTime") & vbCrLf
            Response.Write "�� �� ����" & rsArticleList("Hits") & vbCrLf & "�� �� �֣�" & Mid(rsArticleList("Keyword"), 2, Len(rsArticleList("Keyword")) - 2) & vbCrLf & "�Ƽ��ȼ���"
            If rsArticleList("Stars") = 0 Then
                Response.Write "��"
            Else
                Response.Write String(rsArticleList("Stars"), "��")
            End If
            Response.Write vbCrLf & "��ҳ��ʽ��"
            If rsArticleList("PaginationType") = 0 Then
                Response.Write "����ҳ"
            ElseIf rsArticleList("PaginationType") = 1 Then
                Response.Write "�Զ���ҳ"
            ElseIf rsArticleList("PaginationType") = 2 Then
                Response.Write "�ֶ���ҳ"
            End If
            Response.Write vbCrLf & "�Ķ�������" & rsArticleList("InfoPoint")
            Response.Write "'>" & rsArticleList("title") & "</a>"
            If ManageType = "Receive" And rsArticleList("Receive") = True Then
                Response.Write "&nbsp;"
                
                If rsArticleList("Received") = "" Then
                    NotReceiveUser = rsArticleList("ReceiveUser")
                Else
                    NotReceiveUser = ""
                    arrUser = Split(rsArticleList("ReceiveUser"), ",")
                    For i = 0 To UBound(arrUser)
                        If FoundInArr(rsArticleList("Received"), arrUser(i), "|") = False Then
                            If NotReceiveUser = "" Then
                                NotReceiveUser = arrUser(i)
                            Else
                                NotReceiveUser = NotReceiveUser & "," & arrUser(i)
                            End If
                        End If
                    Next
                End If
                Response.Write "<a href='' onclick='return false' title='"
                Response.Write "Ҫ��ǩ���û���" & rsArticleList("ReceiveUser") & vbCrLf
                Response.Write "�Ѿ�ǩ���û���" & rsArticleList("Received") & vbCrLf
                Response.Write "��δǩ���û���" & NotReceiveUser
                If NotReceiveUser <> "" Then
                    Response.Write "'><font color=red>[ǩ����]</font></a>"
                Else
                    Response.Write "'><font color=green>[��ǩ��]</font></a>"
                End If
            End If
            Response.Write "</td>"
            Response.Write "      <td width='60' align='center'><a href='" & FileName & "&field=Inputer&keyword=" & rsArticleList("Inputer") & "' title='������鿴���û�¼�������" & ChannelShortName & "'>" & rsArticleList("Inputer") & "</a></td>"
               '�޸���˹������
            If ManageType = "PayMoney" Then
                Response.Write "      <td width='60' align='center'>" & rsArticleList("Author") & "</td>"
                Response.Write "      <td width='80' align='center'>" & rsArticleList("Beneficiary") & "</td>"
                Response.Write "      <td width='60' align='center'>" & FormatNumber(rsArticleList("CopyMoney"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                If rsArticleList("Ispayed") = True Then
                    Response.Write "      <td width='40' align='center'><b>��</b></td>"
                Else
                    Response.Write "      <td width='40' align='center'><font color=red><b>��</b></font></td>"
                End If
                If PayStatus Then
                    Response.Write "      <td width='60' align='center'>" & rsArticleList("payDate") & "</td>"
                Else
                    Response.Write "      <td width='60' align='center'>" & rsArticleList("UpdateTime") & "</td>"
                End If
            Else
                Response.Write "      <td width='40' align='center'>" & rsArticleList("Hits") & "</td>"
                Response.Write "      <td width='80' align='center'>"
                If rsArticleList("OnTop") = True Then
                    Response.Write "<font color=blue>��</font> "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsArticleList("Hits") >= HitsOfHot Then
                    Response.Write "<font color=red>��</font> "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsArticleList("Elite") = True Then
                    Response.Write "<font color=green>��</font> "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If Trim(rsArticleList("DefaultPicUrl")) <> "" Then
                    Response.Write "<font color=blue>ͼ</font>"
                Else
                    Response.Write "&nbsp;&nbsp;"
                End If
                If rsArticleList("VoteID") > 0 Then
                    Response.Write "<a href='" & InstallDir & "Vote.asp?ID=" & rsArticleList("VoteID") & "&Action=Show' target='_blank'>��</a>"
                Else
                    Response.Write "&nbsp;&nbsp;"
                End If
                Response.Write "    </td>"
                Response.Write "    <td width='60' align='center'>"
                Select Case rsArticleList("Status")
                Case -2
                    Response.Write "<font color=gray>�˸�</font>"
                Case -1
                    Response.Write "<font color=gray>�ݸ�</font>"
                Case 0
                    Response.Write "<font color=red>�����</font>"
                Case 1
                    Response.Write "<font color=blue>һ��ͨ��</font>"
                Case 2
                    Response.Write "<font color=green>����ͨ��</font>"
                Case 3
                    Response.Write "<font color=black>����ͨ��</font>"
                End Select
                Response.Write "</td>"
            End If
        
            Dim iClassPurview
            If UseCreateHTML > 0 And ObjInstalled_FSO = True And ManageType <> "Special" Then
                Response.Write "    <td width='40' align='center'>"
                If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
                    iClassPurview = ClassPurview
                    ArticlePath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticleList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsArticleList("UpdateTime"), rsArticleList("ArticleID")) & FileExt_Item
                Else
                    iClassPurview = PE_CLng(rsArticleList("ClassPurview"))
                    ArticlePath = HtmlDir & GetItemPath(StructureType, rsArticleList("ParentDir"), rsArticleList("ClassDir"), rsArticleList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsArticleList("UpdateTime"), rsArticleList("ArticleID")) & FileExt_Item
                End If
                If iClassPurview > 0 Or rsArticleList("InfoPoint") > 0 Then
                    Response.Write "<a href='#' title='��Ϊ�������Ķ�Ȩ�ޣ����Բ�������HTML'><font color=green><b>��</b></font></a>"
                Else
                    If fso.FileExists(Server.MapPath(ArticlePath)) Then
                        Response.Write "<a href='#' title='�ļ�λ�ã�" & ArticlePath & "'><b>��</b></a>"
                    Else
                        Response.Write "<font color=red><b>��</b></font>"
                    End If
                End If
                Response.Write "</td>"
            End If
            Select Case ManageType
            Case "Check"
                Response.Write "<td width='120' align='center'>"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsArticleList("Status") <= MyStatus Then
                        If rsArticleList("Status") > -1 Then
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Reject&ArticleID=" & rsArticleList("ArticleID") & "'>ֱ���˸�</a>&nbsp;&nbsp;"
                        End If
                        If rsArticleList("Status") < MyStatus Then
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Check&ArticleID=" & rsArticleList("ArticleID") & "'>���</a>&nbsp;&nbsp;"
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=SetPassed&ArticleID=" & rsArticleList("ArticleID") & "'>ͨ��</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=CancelPassed&ArticleID=" & rsArticleList("ArticleID") & "'>ȡ�����</a>"
                        End If
                    End If
                End If
                Response.Write "</td>"
            Case "PayMoney"
                If rsArticleList("IsPayed") = False And PayStatus = "False" Then
                    Response.Write "<td width='60' align='center'><a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=ConfirmPay&ArticleID=" & rsArticleList("ArticleID") & "'>֧�����</a></td>"
                End If
            Case "HTML"
                Response.Write "    <td width='180' align='left'>&nbsp;"
                If iClassPurview = 0 And rsArticleList("InfoPoint") = 0 And rsArticleList("Status") = 3 And (AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True) Then
                    Response.Write "<a href='Admin_CreateArticle.asp?ChannelID=" & ChannelID & "&Action=CreateArticle&ArticleID=" & rsArticleList("ArticleID") & "' title='���ɱ�" & ChannelShortName & "��HTMLҳ��'>�����ļ�</a>&nbsp;"
                    If fso.FileExists(Server.MapPath(ArticlePath)) Then
                        Response.Write "<a href='" & ArticlePath & "' target='_blank' title='�鿴��" & ChannelShortName & "��HTMLҳ��'>�鿴�ļ�</a>&nbsp;"
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=DelFile&ArticleID=" & rsArticleList("ArticleID") & "' title='ɾ����" & ChannelShortName & "��HTMLҳ��' onclick=""return confirm('ȷ��Ҫɾ����" & ChannelShortName & "��HTMLҳ����');"">ɾ���ļ�</a>&nbsp;"
                    End If
                End If
                Response.Write "</td>"
            Case "Recyclebin"
                Response.Write "<td width='100' align='center'>"
                Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=ConfirmDel&ArticleID=" & rsArticleList("ArticleID") & "' onclick=""return confirm('ȷ��Ҫ����ɾ����" & ChannelShortName & "�𣿳���ɾ�����޷���ԭ��');"">����ɾ��</a> "
                Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Restore&ArticleID=" & rsArticleList("ArticleID") & "'>��ԭ</a>"
                Response.Write "</td>"
            Case "Special"
                Response.Write "<td width='100' align='center'>"
                If rsArticleList("SpecialID") > 0 Then
                    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=DelFromSpecial&InfoID=" & rsArticleList("InfoID") & "' onclick=""return confirm('ȷ��Ҫ����" & ChannelShortName & "��������ר����ɾ����');"">������ר����ɾ��</a> "
                End If
                Response.Write "</td>"
            Case Else
                Response.Write "    <td width='150' align='left'>&nbsp;"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or CheckPurview_Class(arrClass_Input, ParentPath & "," & ClassID) Or UserName = rsArticleList("Inputer") Then
                    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Modify&ArticleID=" & rsArticleList("ArticleID") & "'>�޸�</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>�޸�&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Or UserName = rsArticleList("Inputer") Then
                    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Del&ArticleID=" & rsArticleList("ArticleID") & "' onclick=""return confirm('ȷ��Ҫɾ����" & ChannelShortName & "��ɾ�����㻹���Դӻ���վ�л�ԭ��');"">ɾ��</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>ɾ��&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsArticleList("OnTop") = False Then
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=SetOnTop&ArticleID=" & rsArticleList("ArticleID") & "'>�̶�</a>&nbsp;"
                    Else
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=CancelOnTop&ArticleID=" & rsArticleList("ArticleID") & "'>���</a>&nbsp;"
                    End If
                    If rsArticleList("Elite") = False Then
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=SetElite&ArticleID=" & rsArticleList("ArticleID") & "'>��Ϊ�Ƽ�</a>"
                    Else
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=CancelElite&ArticleID=" & rsArticleList("ArticleID") & "'>ȡ���Ƽ�</a>"
                    End If
                End If
                Response.Write "</td>"
            End Select
            Response.Write "</tr>"

            ArticleNum = ArticleNum + 1
            If ArticleNum >= MaxPerPage Then Exit Do
            rsArticleList.MoveNext
        Loop
    End If
    rsArticleList.Close
    Set rsArticleList = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�б�ҳ��ʾ������" & ChannelShortName & "</td><td>"
    Select Case ManageType
    Case "Check"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='submit1' type='submit' value=' ���ͨ�� ' onClick=""document.myform.Action.value='SetPassed'"">&nbsp;&nbsp;"
            Response.Write "<input name='submit2' type='submit' value=' ȡ����� ' onClick=""document.myform.Action.value='CancelPassed'"">&nbsp;&nbsp;"
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "<input name='submit3' type='submit' value=' ����ɾ�� ' onClick=""document.myform.Action.value='Del'"">"
            End If
        End If
    Case "HTML"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='CreateType' type='hidden' id='CreateType' value='1'>"
            Response.Write "<input name='ClassID' type='hidden' id='ClassID' value='" & ClassID & "'>"
            If ClassID > 0 Then
                If UseCreateHTML = 1 Or UseCreateHTML = 3 And ClassPurview < 2 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='1';document.myform.action='Admin_CreateArticle.asp';"" value='���ɵ�ǰ��Ŀ�б�ҳ'>&nbsp;&nbsp;"
                End If
                If ClassPurview = 0 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateArticle';document.myform.CreateType.value='2';document.myform.action='Admin_CreateArticle.asp';"" value='���ɵ�ǰ��Ŀ��" & ChannelShortName & "'>&nbsp;&nbsp;"
                End If
            Else
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateIndex';document.myform.CreateType.value='1';document.myform.action='Admin_CreateArticle.asp';"" value='������ҳ'>&nbsp;&nbsp;"
                If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='2';document.myform.action='Admin_CreateArticle.asp';"" value='����������Ŀ�б�ҳ'>&nbsp;&nbsp;"
                End If
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateArticle';document.myform.CreateType.value='3';document.myform.action='Admin_CreateArticle.asp';"" value='��������" & ChannelShortName & "'>&nbsp;&nbsp;"
            End If
            Response.Write "<input name='submit2' type='submit' id='submit2' onClick=""document.myform.Action.value='CreateArticle';document.myform.action='Admin_CreateArticle.asp';"" value='����ѡ����" & ChannelShortName & "'>&nbsp;&nbsp;&nbsp;&nbsp;"
            Response.Write "<input name='Submit3' type='submit' id='submit3' onClick=""document.myform.Action.value='DelFile';document.myform.action='Admin_Article.asp'"" value='ɾ��ѡ��" & ChannelShortName & "��HTML�ļ�'>"
        End If
    Case "Recyclebin"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='ConfirmDel'"" value=' ����ɾ�� '>&nbsp;"
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='ClearRecyclebin'"" value='��ջ���վ'>&nbsp;&nbsp;&nbsp;&nbsp;"
            Response.Write "<input name='Submit3' type='submit' id='Submit3' onClick=""document.myform.Action.value='Restore'"" value='��ԭѡ����" & ChannelShortName & "'>&nbsp;"
            Response.Write "<input name='Submit4' type='submit' id='Submit4' onClick=""document.myform.Action.value='RestoreAll'"" value='��ԭ����" & ChannelShortName & "'>"
        End If
    Case "Special"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='DelFromSpecial'"" value='������ר�����Ƴ�'> "
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='AddToSpecial'"" value='��ӵ�����ר����'> "
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='MoveToSpecial'"" value='�ƶ�����һר����'>"
        End If
    Case "PayMoney"
        Response.Write "<Script Language='JavaScript'>"
        Response.Write "function SetBtStatPayValue()"
        Response.Write "{"
        Response.Write "document.myform.Action.value='ConfirmPay';"
        Response.Write "document.myform.submit();"
        Response.Write "}"
        Response.Write "function SetExportExcelValue()"
        Response.Write "{"
        Response.Write "document.myform.Action.value='ExportExcel';"
        Response.Write "document.myform.submit();"
        Response.Write "}"
        Response.Write "</Script>"
        Response.Write "<Input Type='Hidden' Name='ManageType' Value='" & ManageType & "'>"
        Response.Write "<Input Type='Hidden' Name='PayStatus' Value='" & PayStatus & "'>"
        Response.Write "<table border=0>"
        Response.Write "<tr>"
        If PayStatus = "False" Then
            Response.Write "<td><Input name='BtStatPay' type='Button' id='BtStatPay' value='����֧�����' onClick=""SetBtStatPayValue()""></td>"
        End If
        Response.Write "</tr>"
        Response.Write "<tr>"
        Response.Write "<td>"
        Call PopCalendarInit
        Response.Write "<input name='SelectType' type='radio' value='ID' >���ɣķ�Χѡ��"
        Response.Write "��ʼ�ɣ�<input type='text' name='BeginID'  size='10' value='1'>"
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "��ֹ�ɣ�<input type='text' name='EndID'  size='10' value='1000'>"
        Response.Write "&nbsp;&nbsp;&nbsp;<br>"
        Response.Write "<input name='SelectType' type='radio' value='Date'>�����ڷ�Χѡ��"
        Response.Write "��ʼ����<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;��������<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
        If PayStatus = "False" Then
            Response.Write "<input type='button' name='btExportExcel'  value='����δ֧����" & ChannelShortName & "��EXCEL' onClick=""SetExportExcelValue()"">"
            
        Else
            Response.Write "<input type='submit' name='btExportExcel'  value='������֧����" & ChannelShortName & "��EXCEL' onClick=""SetExportExcelValue()"">"
        End If
        
        Response.Write "</td>"
        Response.Write "</tr>"
        Response.Write "</table>"
    Case Else
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='submit1' type='submit' value=' ����ɾ�� ' onClick=""document.myform.Action.value='Del'""> "
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "<input type='submit' name='Submit4' value=' �����ƶ� ' onClick=""document.myform.Action.value='MoveToClass'""> "
                Response.Write "<input type='submit' name='Submit3' value=' �������� ' onClick=""document.myform.Action.value='Batch'""> "
                Response.Write "<input name='submit1' type='submit' value=' ���ͨ�� ' onClick=""document.myform.Action.value='SetPassed'""> "
                Response.Write "<input name='submit2' type='submit' value=' ȡ����� ' onClick=""document.myform.Action.value='CancelPassed'""> "
                Response.Write "<input name='submit3' type='submit' value=' �����滻 ' onClick=""document.myform.Action.value='BatchReplace'""> "
            End If
        End If
    End Select
    
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
'    If SystemDatabaseType = "SQL" Then
'        totalPut = Cmd.Parameters("RETURN_VALUE").Value
'        CurrentPage = Cmd.Parameters("@ActualPage").Value
'    End If
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName & "", True)
    End If

    Response.Write "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    Response.Write "<Input Type='Hidden' Name='PayStatus' Value='" & PayStatus & "'>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>" & ChannelShortName & "������</strong></td>"
    Response.Write "   <td>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='Title' selected>" & ChannelShortName & "����</option>"
    Response.Write "<option value='Content'>" & ChannelShortName & "����</option>"
    Response.Write "<option value='Author'>" & ChannelShortName & "����</option>"
    If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
        Response.Write "<option value='Inputer'>¼����</option>"
        Response.Write "<option value='Editor'>�����</option>"
    End If
    Response.Write "<option value='UpdateTime'>����ʱ��</option>"
    Response.Write "<option value='Keyword'>�ؼ���</option>"
    Response.Write "<option value='ID'>" & ChannelShortName & "ID</option>"
    Response.Write "</select>"
    Response.Write "<select name='ClassID'><option value=''>������Ŀ</option>" & GetClass_Option(1, 0) & "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='����'>"
    Response.Write "<input name='ManageType' type='hidden' id='ManageType' value='" & ManageType & "'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "</td></tr></table></form>"
    Response.Write "<br><b>˵����</b><br>&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "�����еĸ���壺<font color=blue>��</font>----�̶�" & ChannelShortName & "��<font color=red>��</font>----����" & ChannelShortName & "��<font color=green>��</font>----�Ƽ�" & ChannelShortName & "��<font color=blue>ͼ</font>----��ҳͼƬ" & ChannelShortName & "��<font color=black>��</font>----�е���<br><br>"
End Sub

Sub ShowJS_Article()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function AddItem(strFileName){" & vbCrLf
    Response.Write "    var arrName=strFileName.split('.');" & vbCrLf
    Response.Write "    var FileExt=arrName[1];" & vbCrLf
    Response.Write "    if (FileExt=='gif'||FileExt=='jpg'||FileExt=='jpeg'||FileExt=='jpe'||FileExt=='bmp'||FileExt=='png'){" & vbCrLf
    Response.Write "        if (arrName[0].substr(arrName[0].length-2,arrName[0].length)!='_S'){" & vbCrLf
    Response.Write "            if(document.myform.IncludePic.selectedIndex<2){" & vbCrLf
    Response.Write "                document.myform.IncludePic.selectedIndex+=1;" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        frmPreview.img.src='" & InstallDir & ChannelDir & "/" & UploadDir & "/" & "'+strFileName;" & vbCrLf
    Response.Write "        document.myform.DefaultPicUrl.value=strFileName;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);" & vbCrLf
    Response.Write "    document.myform.DefaultPicList.selectedIndex+=1;" & vbCrLf
    Response.Write "    if(document.myform.UploadFiles.value==''){" & vbCrLf
    Response.Write "        document.myform.UploadFiles.value=strFileName;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        document.myform.UploadFiles.value=document.myform.UploadFiles.value+'|'+strFileName;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function selectPaginationType(){" & vbCrLf
    Response.Write "  document.myform.PaginationType.value=2;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function rUseLinkUrl(){" & vbCrLf
    Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=false;" & vbCrLf
    Response.Write "     ArticleContent.style.display='none';" & vbCrLf
    Response.Write "     ArticleContent2.style.display='none';" & vbCrLf
    Response.Write "     ArticleContent3.style.display='none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=true;" & vbCrLf
    Response.Write "    ArticleContent.style.display='';" & vbCrLf
    Response.Write "    ArticleContent2.style.display='';" & vbCrLf
    Response.Write "    ArticleContent3.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf

    Response.Write "  else{" & vbCrLf
    Response.Write "    alert('Ԥ��״̬���ܱ��棡���Ȼص��༭״̬���ٱ���');" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf

    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "���ⲻ��Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.Title.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('�ؼ��ֲ���Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
    Response.Write "    if (document.myform.LinkUrl.value==''||document.myform.LinkUrl.value=='http://'){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('������ת�����ӵĵ�ַ��');" & vbCrLf
    Response.Write "      document.myform.LinkUrl.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('" & ChannelShortName & "���ݲ���Ϊ�գ�');" & vbCrLf
    Response.Write "      editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  var obj=document.myform.ClassID;" & vbCrLf
    Response.Write "  var iCount=0;" & vbCrLf
    Response.Write "  for(var i=0;i<obj.length;i++){" & vbCrLf
    Response.Write "    if(obj.options[i].selected==true){" & vbCrLf
    Response.Write "      iCount=iCount+1;" & vbCrLf
    Response.Write "      if(obj.options[i].value==''){" & vbCrLf
    Response.Write "        ShowTabs(0);" & vbCrLf
    Response.Write "        alert('" & ChannelShortName & "������Ŀ����ָ��Ϊ�ⲿ��Ŀ��');" & vbCrLf
    Response.Write "        document.myform.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "      if(obj.options[i].selected==true&&obj.options[i].value=='0'){" & vbCrLf
    Response.Write "        ShowTabs(0);" & vbCrLf
    Response.Write "        alert('ָ������Ŀ���������" & ChannelShortName & "��ֻ������������Ŀ�����" & ChannelShortName & "��');" & vbCrLf
    Response.Write "        document.myform.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (iCount==0){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('��ѡ��������Ŀ��');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID>0){" & vbCrLf
    Response.Write "    Tabs_Bottom.style.display='none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    Tabs_Bottom.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "    if(ID==0){" & vbCrLf
    Response.Write "      editor.yToolbarsCss();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "function SelectUser(){" & vbCrLf
    Response.Write "    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=UserList&DefaultValue='+document.myform.InceptUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        document.myform.InceptUser.value=arr;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "function getPayMoney(){" & vbCrLf
    'Response.Write "alert(document.myform.PerWordMoney.value*document.myform.WordNumber.value/1000);"
    Response.Write "  document.myform.CopyMoney1.value=document.myform.PerWordMoney.value*document.myform.WordNumber.value/1000;" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "function IsDigit(){" & vbCrLf
    Response.Write "  return ((event.keyCode >= 48) && (event.keyCode <= 57));" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function CopyTitle(){" & vbCrLf
    Response.Write "  if (document.myform.VoteTitle.value==''){" & vbCrLf
    Response.Write "     document.myform.VoteTitle.value = document.myform.Title.value;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function moreitem(inputname,listnum,ichannelid,inputype){" & vbCrLf
    Response.Write "    var chedkurl = '../inc/checklist.asp';" & vbCrLf
    Response.Write "    var CheckDOM = new ActiveXObject(""Microsoft.XMLDOM"");" & vbCrLf
    Response.Write "    CheckDOM.async=false;" & vbCrLf
    Response.Write "    var p = CheckDOM.createProcessingInstruction('xml','version=\""1.0\"" encoding=\""gb2312\""'); " & vbCrLf
    Response.Write "    CheckDOM.appendChild(p); " & vbCrLf

    Response.Write "    var CheckRoot = CheckDOM.createElement('root');" & vbCrLf
    Response.Write "    var CField = CheckDOM.createNode(1,'text',''); " & vbCrLf
    Response.Write "    CField.text = $F(inputname);" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'lnum',''); " & vbCrLf
    Response.Write "    CField.text = listnum;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'channelid',''); " & vbCrLf
    Response.Write "    CField.text = ichannelid;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'type',''); " & vbCrLf
    Response.Write "    CField.text = inputype;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'inputname',''); " & vbCrLf
    Response.Write "    CField.text = inputname;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CheckDOM.appendChild(CheckRoot);" & vbCrLf

    Response.Write "    var CHttp = getHTTPObject();" & vbCrLf
    Response.Write "    CHttp.open('POST',chedkurl,true);" & vbCrLf
    Response.Write "    CHttp.onreadystatechange = function () " & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(CHttp.readyState == 4 && CHttp.status==200){" & vbCrLf
    Response.Write "            if(CHttp.responseText == ''){" & vbCrLf
    Response.Write "                Element.hide(inputype);" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                Element.show(inputype);" & vbCrLf
    Response.Write "                $(inputype).innerHTML=CHttp.responseText;" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    CHttp.send(CheckDOM);" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function addinput(iname,ivalue){" & vbCrLf
    Response.Write "  if(iname!='' && ivalue!=''){" & vbCrLf
    Response.Write "     $(iname).value=ivalue;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    'Response.Write "function getKey() {" & vbCrLf
    'Response.Write " if(window.event.keyCode==49) {" & vbCrLf
    'Response.Write "   ShowTabs(0);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==50) {" & vbCrLf
    'Response.Write "   ShowTabs(1);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==51) {" & vbCrLf
    'Response.Write "   ShowTabs(2);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==52) {" & vbCrLf
    'Response.Write "   ShowTabs(3);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==53) {" & vbCrLf
    'Response.Write "   ShowTabs(5);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==54) {" & vbCrLf
    'Response.Write "   ShowTabs(6);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==55) {" & vbCrLf
    'Response.Write "   ShowTabs(7);CopyTitle();" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==56) {" & vbCrLf
    'Response.Write "   ShowTabs(4);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write "}" & vbCrLf
    'Response.Write "document.onkeypress = getKey;" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Function ReplaceJSImage(ByVal Content)
    If Content = "" Then
        ReplaceJSImage = Content
        Exit Function
    End If
    Dim strTemp
    'ͼƬ�滻JS
    regEx.Pattern = "(\<Script)(.[^\<]*)(\<\/Script\>)"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        strTemp = Replace(Match.value, "<", "[!")
        strTemp = Replace(strTemp, ">", "!]")
        strTemp = Replace(strTemp, "'", """")
        strTemp = "<IMG alt='#" & strTemp & "#' src=""" & InstallDir & "editor/images/jscript.gif"" border=0 $>"
        Content = Replace(Content, Match.value, strTemp)
    Next
    ReplaceJSImage = Content
End Function

Sub ShowTabs_Title()
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>����ר��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>��������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'"
    If FoundInArr(arrEnabledTabs, "Charge", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'"
    If FoundInArr(arrEnabledTabs, "Recieve", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">ǩ������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'"
    If FoundInArr(arrEnabledTabs, "Copyfee", ",") = False Or Action = "Add" Then Response.Write " style='display:none'"
    Response.Write ">�������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(6);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">��������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(7)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub ShowTabs_Bottom()
    Response.Write "<table id='Tabs_Bottom' width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title4' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(1)'>����ר��</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(2)'>��������</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(3)'"
    If FoundInArr(arrEnabledTabs, "Charge", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(4)'"
    If FoundInArr(arrEnabledTabs, "Recieve", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">ǩ������</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(5)'"
    If FoundInArr(arrEnabledTabs, "Copyfee", ",") = False Or Action = "Add" Then Response.Write " style='display:none'"
    Response.Write ">�������</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(6);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">��������</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(7)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Add()
    Call ShowJS_Article
    
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "����</a>&nbsp;&gt;&gt;&nbsp;���" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Article.asp' target='_self'>"

    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���⣺</td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='2'>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64' class='tdbg5'>��̱��⣺</td>"
    Response.Write "                  <td>"
    
    Response.Write "                    <select name='IncludePic'>"
    Response.Write "                      <option  value='0' selected> </option>"
    Response.Write "                      <option value='1'>" & ArticlePro1 & "</option>"
    Response.Write "                      <option value='2'>" & ArticlePro2 & "</option>"
    Response.Write "                      <option value='3'>" & ArticlePro3 & "</option>"
    Response.Write "                      <option value='4'>" & ArticlePro4 & "</option>"
    Response.Write "                    </select>"
    
   
    Response.Write "                    <input name='Title' type='text' id='Title' value='' size='56' autocomplete='off' maxlength='255' class='bginput' onPropertyChange=""moreitem('Title',10," & ChannelID & ",'satitle');"" onBlur=""setTimeout('Element.hide(satitle)',500);"">"
    Response.Write "                    <select name='TitleFontColor' id='TitleFontColor'>"
    Response.Write "                      <option value='' selected>��ɫ</option>"
    Response.Write "                      <option value=''>Ĭ��</option>"
    Response.Write "                      <option value='#000000' style='background-color:#000000'></option>"
    Response.Write "                      <option value='#FFFFFF' style='background-color:#FFFFFF'></option>"
    Response.Write "                      <option value='#008000' style='background-color:#008000'></option>"
    Response.Write "                      <option value='#800000' style='background-color:#800000'></option>"
    Response.Write "                      <option value='#808000' style='background-color:#808000'></option>"
    Response.Write "                      <option value='#000080' style='background-color:#000080'></option>"
    Response.Write "                      <option value='#800080' style='background-color:#800080'></option>"
    Response.Write "                      <option value='#808080' style='background-color:#808080'></option>"
    Response.Write "                      <option value='#FFFF00' style='background-color:#FFFF00'></option>"
    Response.Write "                      <option value='#00FF00' style='background-color:#00FF00'></option>"
    Response.Write "                      <option value='#00FFFF' style='background-color:#00FFFF'></option>"
    Response.Write "                      <option value='#FF00FF' style='background-color:#FF00FF'></option>"
    Response.Write "                      <option value='#FF0000' style='background-color:#FF0000'></option>"
    Response.Write "                      <option value='#0000FF' style='background-color:#0000FF'></option>"
    Response.Write "                      <option value='#008080' style='background-color:#008080'></option>"
    Response.Write "                    </select>"
    Response.Write "                    <select name='TitleFontType' id='TitleFontType'>"
    Response.Write "                      <option value='0' selected>����</option>"
    Response.Write "                      <option value='1'>����</option>"
    Response.Write "                      <option value='2'>б��</option>"
    Response.Write "                      <option value='3'>��+б</option>"
    Response.Write "                      <option value='0'>����</option>"
    Response.Write "                    </select>"
    Response.Write "                    <div id=""satitle"" style='display:none'></div>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64' class='tdbg5'>�������⣺</td>"
    Response.Write "                  <td><input name='TitleIntact' type='text' id='TitleIntact' value='' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64' class='tdbg5'>�� �� �⣺</td>"
    Response.Write "                  <td><input name='Subheading' type='text' id='Subheading' value='' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td></td><td><input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'>��ʾ" & ChannelShortName & "�б�ʱ�ڱ�������ʾ��������&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='checksame' value='����ظ�����' onclick=""showModalDialog('Admin_CheckSameTitle.asp?ModuleType=" & ModuleType & "&Title='+document.myform.Title.value,'checksame','dialogWidth:350px; dialogHeight:250px; help: no; scroll: no; status: no');""></td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ؼ��֣�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' style=""clear:both"" id='Keyword' value='" & Trim(Session("Keyword")) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div><font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���ߣ�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('Author',10," & ChannelID & ",'sauthor');"" onBlur=""setTimeout('Element.hide(sauthor)',500);"">" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "              </div><div id=""sauthor"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��Դ��</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('CopyFrom',10," & ChannelID & ",'scopyfrom');"" onBlur=""setTimeout('Element.hide(scopyfrom)',500);"">" & GetCopyFromList("Admin", ChannelID)
    Response.Write "              </div><div id=""scopyfrom"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><font color='#FF0000'>ת�����ӣ�</font></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='LinkUrl' type='text' id='LinkUrl' value='http://' size='50' maxlength='255' disabled>"
    Response.Write "              <input name='UseLinkUrl' type='checkbox' id='UseLinkUrl' value='Yes' onClick='rUseLinkUrl();'> <font color='#FF0000'>ʹ��ת������</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��飺</td>"
    Response.Write "            <td><textarea name='Intro' cols='80' rows='4'></textarea></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent' style=""display:''"">"
    Response.Write "            <td width='120' align='right' valign='bottom' class='tdbg5'><p>" & ChannelShortName & "���ݣ�</p>"
    If EnableSaveRemote = True And IsObjInstalled("Microsoft.XMLHTTP") = True Then
        Response.Write "<table><tr><td><input type='checkbox' name='SaveRemotePic' value='Yes' checked></td><td>�Զ�����" & ChannelShortName & "�������ͼƬ</td>"
        If PhotoObject = 1 Then
            Response.Write "<tr><td><input type='checkbox' name='AddWatermark' value='Yes' checked></td><td>�Ƿ��" & ChannelShortName & "������Զ�̻�õ�ͼƬ��ˮӡ</td></tr>"
            Response.Write "<tr><td><input type='checkbox' name='AddThumb' value='Yes' checked></td><td>�Ƿ��" & ChannelShortName & "������Զ�̻�õĵ�һ��ͼƬ������ͼ</td></tr>"
        End If
        Response.Write "</tr></table>"
        Response.Write "<div align='left'><font color='#006600'>&nbsp;&nbsp;&nbsp;&nbsp;���ô˹��ܺ������������վ�ϸ������ݵ��ұߵı༭���У����������а�����ͼƬ����ϵͳ���ڱ���" & ChannelShortName & "ʱ�Զ������ͼƬ���Ƶ���վ�������ϡ�"
        Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;ϵͳ����������ͼƬ�Ĵ�С��Ӱ���ٶȣ�����ͼƬ�϶�ʱ��Ҫʹ�ô˹��ܡ�</font>"
    End If
    Response.Write "<br><br><font color='red'>�����밴Shift+Enter<br><br>����һ���밴Enter</font></div><br><br><br><br><iframe id='frmPreview' width='120' height='150' frameborder='1' src='Admin_imgPreview.asp'></iframe>"
    Response.Write "            </td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='Content' style='display:none'>" & XmlText("Article", "DefaultAddTemplate", "") & "</textarea>"
    Response.Write "              <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content' frameborder='1' scrolling='no' width='600' height='600' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><font color='#FF0000'>��ҳͼƬ��</font></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='DefaultPicUrl' type='text' id='DefaultPicUrl' size='56' maxlength='200'>"
    Response.Write "              ��������ҳ��ͼƬ" & ChannelShortName & "����ʾ <br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��"
    Response.Write "              <select name='DefaultPicList' id='DefaultPicList' onChange=""DefaultPicUrl.value=this.value;frmPreview.img.src=((this.value == '') ? '../images/nopic.gif' : '" & InstallDir & ChannelDir & "/" & UploadDir & "/'+this.value);"">"
    Response.Write "                <option selected>��ָ����ҳͼƬ</option>"
    Response.Write "              </select>"
    Response.Write "              <input name='UploadFiles' type='hidden' id='UploadFiles'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent2' style=""display:''""> "
    Response.Write "            <td width='120' align='right' class='tdbg5'>���ݷ�ҳ��ʽ��</td>"
    Response.Write "            <td>"
    Response.Write "              <select name='PaginationType' id='PaginationType'>"
    Response.Write "                <option value='0' selected>����ҳ</option>"
    Response.Write "                <option value='1'>�Զ���ҳ</option>"
    Response.Write "                <option value='2'>�ֶ���ҳ</option>"
    Response.Write "              </select>"
    Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;<strong><font color='#0000FF'>ע��</font></strong><font color='#0000FF'>�ֶ���ҳ�����Ϊ��</font><font color='#FF0000'>[NextPage]</font><font color='#0000FF'>����ע���Сд</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent3' style=""display:''"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'>&nbsp;</td>"
    Response.Write "            <td>�Զ���ҳʱ��ÿҳ��Լ�ַ���������HTML����ұ������100����<input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value='10000' size='8' maxlength='8'></td>"
    Response.Write "          </tr>"
    
    Call ShowTabs_Status_Add
    
    Response.Write "        </tbody>" & vbCrLf
    
    
    Call ShowTabs_Special(SpecialID, "")

    Call ShowTabs_Property_Add
    
    Call ShowTabs_Purview_Add("�Ķ�")
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>ǩ���û���</td>"
    Response.Write "            <td><textarea name='InceptUser' cols='72' rows='5' readonly></textarea><br>"
    Response.Write "                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_choose' value='ѡ���û�' onClick='SelectUser();'>"
    Response.Write "                    &nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_cancel' value='����û�' onClick=""myform.InceptUser.value=''"">"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>ǩ�շ�ʽ��</td>"
    Response.Write "            <td>"
    Response.Write "                    <select name='AutoReceiveTime'>"
    Response.Write "                      <option value='0' selected>�ֶ�ǩ��</option>"
    Response.Write "                      <option value='5'>5���Ӻ�</option>"
    Response.Write "                      <option value='10'>10���Ӻ�</option>"
    Response.Write "                      <option value='30'>30���Ӻ�</option>"
    Response.Write "                      <option value='60'>1���Ӻ�</option>"
    Response.Write "                      <option value='120'>2���Ӻ�</option>"
    Response.Write "                      <option value='300'>5���Ӻ�</option>"
    Response.Write "                    </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ĵ����ͣ�</td>"
    Response.Write "            <td><select name='ReceiveType' id='ReceiveType'><option value='0' selected>�����ĵ�</option><option value='1'>ר���ĵ�</option></select></td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf

    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
   
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Vote_Add
    
    Call ShowTabs_MyField_Add
    
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"

    Call ShowTabs_Bottom
    
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' �� �� ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp; "
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' Ԥ �� ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub


Sub Modify()
    Dim rsArticle, sql, tmpAuthor, tmpCopyFrom

    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        ArticleID = PE_CLng(ArticleID)
    End If
    sql = "select * from PE_Article where ArticleID=" & ArticleID & ""
    Set rsArticle = Conn.Execute(sql)
    If rsArticle.BOF And rsArticle.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "</li>"
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If

    ClassID = rsArticle("ClassID")
    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, ArticleID)

    If rsArticle("Inputer") <> UserName Then
        Call CheckClassPurview(Action, ClassID)
    End If
    If FoundErr = True Then
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If
    tmpAuthor = rsArticle("Author")
    tmpCopyFrom = rsArticle("CopyFrom")
    EmailOfReject = Replace(EmailOfReject, "{$Title}", rsArticle("Title"))
    EmailOfPassed = Replace(EmailOfPassed, "{$Title}", rsArticle("Title"))

    Call ShowJS_Article


    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "����</a>&nbsp;&gt;&gt;&nbsp;�޸�" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Article.asp' target='_self'>"

    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���⣺</td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%'  border='0' cellspacing='2' cellpadding='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>��̱��⣺</td>"
    Response.Write "                  <td>"

    Response.Write "                    <select name='IncludePic'>"
    Response.Write "                      <option  value='0'"
    If rsArticle("IncludePic") = 0 Then Response.Write " selected"
    Response.Write "> </option>"
    Response.Write "                      <option value='1'"
    If rsArticle("IncludePic") = 1 Then Response.Write " selected"
    Response.Write ">" & ArticlePro1 & "</option>"
    Response.Write "                      <option value='2'"
    If rsArticle("IncludePic") = 2 Then Response.Write " selected"
    Response.Write ">" & ArticlePro2 & "</option>"
    Response.Write "                      <option value='3'"
    If rsArticle("IncludePic") = 3 Then Response.Write " selected"
    Response.Write ">" & ArticlePro3 & "</option>"
    Response.Write "                      <option value='4'"
    If rsArticle("IncludePic") = 4 Then Response.Write " selected"
    Response.Write ">" & ArticlePro4 & "</option>"
    Response.Write "                    </select>"
        
    Response.Write "                    <input name='Title' type='text' id='Title' value='" & rsArticle("Title") & "' autocomplete='off' size='56' maxlength='255' class='bginput' onPropertyChange=""moreitem('Title',10," & ChannelID & ",'satitle');"" onBlur=""setTimeout('Element.hide(satitle)',500);"">"
    Response.Write "                    <select name='TitleFontColor' id='TitleFontColor'>"
    Response.Write "                      <option value=''"
    If rsArticle("TitleFontColor") = "" Then Response.Write " selected"
    Response.Write ">��ɫ</option>"
    Response.Write "                      <option value=''>Ĭ��</option>"
    Response.Write "                      <option value='#000000' style='background-color:#000000'"
    If rsArticle("TitleFontColor") = "#000000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FFFFFF' style='background-color:#FFFFFF'"
    If rsArticle("TitleFontColor") = "#FFFFFF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#008000' style='background-color:#008000'"
    If rsArticle("TitleFontColor") = "#008000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#800000' style='background-color:#800000'"
    If rsArticle("TitleFontColor") = "#800000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#808000' style='background-color:#808000'"
    If rsArticle("TitleFontColor") = "#808000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#000080' style='background-color:#000080'"
    If rsArticle("TitleFontColor") = "#000080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#800080' style='background-color:#800080'"
    If rsArticle("TitleFontColor") = "#800080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#808080' style='background-color:#808080'"
    If rsArticle("TitleFontColor") = "#808080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FFFF00' style='background-color:#FFFF00'"
    If rsArticle("TitleFontColor") = "#FFFF00" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#00FF00' style='background-color:#00FF00'"
    If rsArticle("TitleFontColor") = "#00FF00" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#00FFFF' style='background-color:#00FFFF'"
    If rsArticle("TitleFontColor") = "#00FFFF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FF00FF' style='background-color:#FF00FF'"
    If rsArticle("TitleFontColor") = "#FF00FF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FF0000' style='background-color:#FF0000'"
    If rsArticle("TitleFontColor") = "#FF0000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#0000FF' style='background-color:#0000FF'"
    If rsArticle("TitleFontColor") = "#0000FF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#008080' style='background-color:#008080'"
    If rsArticle("TitleFontColor") = "#008080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                    </select>"
    Response.Write "                    <select name='TitleFontType' id='TitleFontType'>"
    Response.Write "                      <option value='0'"
    If rsArticle("TitleFontType") = 0 Then Response.Write " selected"
    Response.Write ">����</option>"
    Response.Write "                      <option value='1'"
    If rsArticle("TitleFontType") = 1 Then Response.Write " selected"
    Response.Write ">����</option>"
    Response.Write "                      <option value='2'"
    If rsArticle("TitleFontType") = 2 Then Response.Write " selected"
    Response.Write ">б��</option>"
    Response.Write "                      <option value='3'"
    If rsArticle("TitleFontType") = 3 Then Response.Write " selected"
    Response.Write ">��+б</option>"
    Response.Write "                      <option value='0'"
    If rsArticle("TitleFontType") = 4 Then Response.Write " selected"
    Response.Write ">����</option>"
    Response.Write "                    </select>"
    Response.Write "                    <div id=""satitle"" style='display:none'></div>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>�������⣺</td>"
    Response.Write "                  <td><input name='TitleIntact' type='text' id='TitleIntact' value='" & rsArticle("TitleIntact") & "' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>�� �� �⣺</td>"
    Response.Write "                  <td><input name='Subheading' type='text' id='Subheading' value='" & rsArticle("Subheading") & "' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>&nbsp;</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'"
    If rsArticle("ShowCommentLink") = True Then Response.Write "checked"
    Response.Write ">��ʾ" & ChannelShortName & "�б�ʱ�ڱ�������ʾ��������</td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ؼ��֣�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' id='Keyword' value='" & Mid(rsArticle("Keyword"), 2, Len(rsArticle("Keyword")) - 2) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div>"
    Response.Write "              <font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���ߣ�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Author' type='text' id='Author' value='" & tmpAuthor & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('Author',10," & ChannelID & ",'sauthor');"" onBlur=""setTimeout('Element.hide(sauthor)',500);"">" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "              </div><div id=""sauthor"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��Դ��</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('CopyFrom',10," & ChannelID & ",'scopyfrom');"" onBlur=""setTimeout('Element.hide(scopyfrom)',500);"">" & GetCopyFromList("Admin", ChannelID)
    Response.Write "              </div><div id=""scopyfrom"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><font color='#FF0000'>ת�����ӣ�</font></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='LinkUrl' type='text' id='LinkUrl' value='" & rsArticle("LinkUrl") & "' size='50' maxlength='255'"
    If rsArticle("LinkUrl") = "" Or rsArticle("LinkUrl") = "http://" Then Response.Write " disabled"
    Response.Write "> <input name='UseLinkUrl' type='checkbox' id='UseLinkUrl' value='Yes' onClick='rUseLinkUrl();'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write " checked"
    Response.Write "><font color='#FF0000'>ʹ��ת������</font></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��飺</td>"
    Response.Write "            <td><textarea name='Intro' cols='80' rows='4'>" & rsArticle("Intro") & "</textarea></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent' style=""display:'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write "none"
    Response.Write "'"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'><p>" & ChannelShortName & "���ݣ�</p>"
    If EnableSaveRemote = True And IsObjInstalled("Microsoft.XMLHTTP") = True Then
        Response.Write "<table><tr><td><input type='checkbox' name='SaveRemotePic' value='Yes' checked></td><td>�Զ�����" & ChannelShortName & "�������ͼƬ</td>"
        If PhotoObject = 1 Then
            Response.Write "<tr><td><input type='checkbox' name='AddWatermark' value='Yes' checked></td><td>�Ƿ��" & ChannelShortName & "������Զ�̻�õ�ͼƬ��ˮӡ</td></tr>"
            Response.Write "<tr><td><input type='checkbox' name='AddThumb' value='Yes' checked></td><td>�Ƿ��" & ChannelShortName & "������Զ�̻�õĵ�һ��ͼƬ������ͼ</td></tr>"
        End If
        Response.Write "</table>"
        Response.Write "<div align='left'><font color='#006600'>&nbsp;&nbsp;&nbsp;&nbsp;���ô˹��ܺ������������վ�ϸ������ݵ��ұߵı༭���У����������а�����ͼƬ����ϵͳ���ڱ���" & ChannelShortName & "ʱ�Զ������ͼƬ���Ƶ���վ�������ϡ�"
        Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;ϵͳ����������ͼƬ�Ĵ�С��Ӱ���ٶȣ�����ͼƬ�϶�ʱ��Ҫʹ�ô˹��ܡ�</font>"
    End If
    Response.Write "<br><br><font color='red'>�����밴Shift+Enter<br><br>����һ���밴Enter</font></div><br><br><br><br><iframe id='frmPreview' width='120' height='150' frameborder='1' src='Admin_imgPreview.asp'></iframe>"
    Response.Write "            </td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='Content' style='display:none'>" & ReplaceJSImage(Replace(Replace(Server.HTMLEncode(FilterBadTag(rsArticle("Content"), rsArticle("Inputer"))), "[InstallDir_ChannelDir]", InstallDir & ChannelDir & "/"), "{$UploadDir}", UploadDir)) & "</textarea>"
    Response.Write "              <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content' frameborder='1' scrolling='no' width='600' height='600' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ҳͼƬ��</td>"
    Response.Write "            <td><input name='DefaultPicUrl' type='text' id='DefaultPicUrl' value='" & rsArticle("DefaultPicUrl") & "' size='56' maxlength='200'>"
    Response.Write "              ��������ҳ��ͼƬ" & ChannelShortName & "����ʾ <br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��"
    Response.Write "              <select name='DefaultPicList' id='DefaultPicList' onChange=""DefaultPicUrl.value=this.value;frmPreview.img.src=((this.value == '') ? '../images/nopic.gif' : '" & InstallDir & ChannelDir & "/" & UploadDir & "/'+this.value);"">"
    Response.Write "                <option value=''"
    If rsArticle("DefaultPicUrl") = "" Then Response.Write "selected"
    Response.Write ">��ָ����ҳͼƬ</option>"
    If rsArticle("UploadFiles") <> "" Then
        Dim IsOtherUrl
        IsOtherUrl = True
        If InStr(rsArticle("UploadFiles"), "|") > 1 Then
            Dim arrUploadFiles, intTemp
            arrUploadFiles = Split(rsArticle("UploadFiles"), "|")
            For intTemp = 0 To UBound(arrUploadFiles)
                If rsArticle("DefaultPicUrl") = arrUploadFiles(intTemp) Then
                    Response.Write "<option value='" & arrUploadFiles(intTemp) & "' selected>" & arrUploadFiles(intTemp) & "</option>"
                    IsOtherUrl = False
                Else
                    Response.Write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
                End If
            Next
        Else
            If rsArticle("UploadFiles") = rsArticle("DefaultPicUrl") Then
                Response.Write "<option value='" & rsArticle("UploadFiles") & "' selected>" & rsArticle("UploadFiles") & "</option>"
                IsOtherUrl = False
            Else
                Response.Write "<option value='" & rsArticle("UploadFiles") & "'>" & rsArticle("UploadFiles") & "</option>"
            End If
        End If
        If IsOtherUrl = True And rsArticle("DefaultPicUrl") <> "" Then
            Response.Write "<option value='" & rsArticle("DefaultPicUrl") & "' selected>" & rsArticle("DefaultPicUrl") & "</option>"
        End If
    End If
    Response.Write "              </select>"
    Response.Write "              <input name='UploadFiles' type='hidden' id='UploadFiles' value='" & rsArticle("UploadFiles") & "'> "
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent2' style=""display:'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write "none"
    Response.Write "'"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'>���ݷ�ҳ��ʽ��</td>"
    Response.Write "            <td>"
    Response.Write "              <select name='PaginationType' id='PaginationType'>"
    Response.Write "                <option value='0'"
    If rsArticle("PaginationType") = 0 Then Response.Write " selected"
    Response.Write ">����ҳ</option>"
    Response.Write "                <option value='1'"
    If rsArticle("PaginationType") = 1 Then Response.Write " selected"
    Response.Write ">�Զ���ҳ</option>"
    Response.Write "                <option value='2'"
    If rsArticle("PaginationType") = 2 Then Response.Write " selected"
    Response.Write ">�ֶ���ҳ</option>"
    Response.Write "              </select>"
    Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;<strong><font color='#0000FF'>ע��</font></strong><font color='#0000FF'>�ֶ���ҳ�����Ϊ��</font><font color='#FF0000'>[NextPage]</font><font color='#0000FF'>����ע���Сд</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent3' style=""display:'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write "none"
    Response.Write "'"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'>&nbsp;</td>"
    Response.Write "            <td>�Զ���ҳʱ��ÿҳ��Լ�ַ���������HTML����ұ������100����<input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value='" & rsArticle("MaxCharPerPage") & "' size='8' maxlength='8'></td>"
    Response.Write "          </tr>"
    Call ShowTabs_Status_Modify(rsArticle)
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Special(arrSpecialID, "")
    
    Call ShowTabs_Property_Modify(rsArticle)
    
    Call ShowTabs_Purview_Modify("�Ķ�", rsArticle, "")
    
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>ǩ���û���</td>"
    Response.Write "            <td><textarea name='InceptUser' cols='72' rows='3' readonly>" & rsArticle("ReceiveUser") & "</textarea><br>"
    Response.Write "                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_choose' value='ѡ���û�' onClick='SelectUser();'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_cancel' value='����û�' onClick=""myform.InceptUser.value=''""></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>ǩ�շ�ʽ��</td>"
    Response.Write "            <td><select name='AutoReceiveTime'>"
    Response.Write "                      <option value='0'"
    If rsArticle("AutoReceiveTime") = "0" Then Response.Write " selected"
    Response.Write ">�ֶ�ǩ��</option>"
    Response.Write "                      <option value='5'"
    If rsArticle("AutoReceiveTime") = "5" Then Response.Write " selected"
    Response.Write ">5���Ӻ�</option>"
    Response.Write "                      <option value='10'"
    If rsArticle("AutoReceiveTime") = "10" Then Response.Write " selected"
    Response.Write ">10���Ӻ�</option>"
    Response.Write "                      <option value='30'"
    If rsArticle("AutoReceiveTime") = "30" Then Response.Write " selected"
    Response.Write ">30���Ӻ�</option>"
    Response.Write "                      <option value='60'"
    If rsArticle("AutoReceiveTime") = "60" Then Response.Write " selected"
    Response.Write ">1���Ӻ�</option>"
    Response.Write "                      <option value='120'"
    If rsArticle("AutoReceiveTime") = "120" Then Response.Write " selected"
    Response.Write ">2���Ӻ�</option>"
    Response.Write "                      <option value='300'"
    If rsArticle("AutoReceiveTime") = "300" Then Response.Write " selected"
    Response.Write ">5���Ӻ�</option>"
    Response.Write "                    </select>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ĵ����ͣ�</td>"
    Response.Write "            <td><select name='ReceiveType' id='ReceiveType'>"
    Response.Write "                      <option value='0'"
    If rsArticle("ReceiveType") = "0" Then Response.Write " selected"
    Response.Write ">�����ĵ�</option>"
    Response.Write "                      <OPTION value='1'"
    If rsArticle("ReceiveType") = "1" Then Response.Write " selected"
    Response.Write ">ר���ĵ�</OPTION>"
    Response.Write "                    </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Dim WordNum, strArticle, strSql, rsChannel, payNum
    '���˿ո�HTML�ַ�����������
    WordNum = getWordNumber(rsArticle("Content"))
    
    If MoneyPerKw <= 0 Then
       payNum = 0
    Else
       payNum = MoneyPerKw / 1000 * WordNum
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>����������</td><td><input type='text' name='WordNumber' MaxLength='10' size=6 disabled value='" & WordNum & "'>&nbsp;��</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>֧����׼��</td><td><input type='text' name='PerWordMoney' MaxLength='10'size=6 value=' " & MoneyPerKw & "'ONKEYPRESS=""event.returnValue=IsDigit();"">&nbsp;Ԫ/ǧ��</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>���Ƹ�ѣ�</td><td><input type ='text' name='CopyMoney1' MaxLength='10' size='6' value='" & payNum & "' disabled> Ԫ&nbsp;&nbsp;&nbsp;&nbsp;<input type=button name=payCalculate value=' ���� ' onclick='getPayMoney();'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>ȷ��֧����ѣ�</td><td><input type ='text' name='CopyMoney' ONKEYPRESS=""event.returnValue=IsDigit();"" MaxLength='10' size='6'"
    If rsArticle("IsPayed") = True Then
        Response.Write " disabled"
    End If
    Response.Write " value='" & rsArticle("CopyMoney") & "'> Ԫ</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��������ߣ�</td><td><input type='text' name='Beneficiary' size='20' value='" & rsArticle("Inputer") & "'>&nbsp;&nbsp;<font color=blue>���������֮����"",""����</font></td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    
    Call ShowTabs_Vote_Modify(rsArticle)

    Call ShowTabs_MyField_Modify(rsArticle)
        
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>" & vbCrLf

    Call ShowTabs_Bottom

    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "   <input name='ArticleID' type='hidden' id='ArticleID' value='" & rsArticle("ArticleID") & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='�����޸Ľ��' onClick=""document.myform.Action.value='SaveModify';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Save' type='submit' value='���Ϊ��" & ChannelShortName & "' onClick=""document.myform.Action.value='SaveModifyAsAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' Ԥ �� ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'><br>"
    Response.Write "  </p><br>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "<script language='javascript'>setTimeout('setpic()',1000);" & vbCrLf
    Response.Write "function setpic(){" & vbCrLf
    If rsArticle("DefaultPicUrl") <> "" Then
        If Left(rsArticle("DefaultPicUrl"), 1) <> "/" And InStr(rsArticle("DefaultPicUrl"), "://") <= 0 Then
            Response.Write "frmPreview.img.src='" & InstallDir & ChannelDir & "/" & UploadDir & "/" & rsArticle("DefaultPicUrl") & "';"
        Else
            Response.Write "frmPreview.img.src='" & rsArticle("DefaultPicUrl") & "';"
        End If
    End If
    Response.Write "}" & vbCrLf
    Response.Write "</script>"
    
    rsArticle.Close
    Set rsArticle = Nothing


End Sub

Sub SaveArticle()
    Dim rsArticle, sql, trs, i
    Dim ArticleID, ClassID, SpecialID, Title, Content
    Dim Keyword, Author, tAuthor, CopyFrom, Inputer, Editor, UpdateTime
    Dim arrUploadFiles, LinkUrl, UseLinkUrl
    Dim ReceiveUser
    Dim arrSpecialID

    ArticleID = Trim(Request.Form("ArticleID"))
    ClassID = Trim(Request.Form("ClassID"))
    SpecialID = Trim(Request.Form("SpecialID"))

    Title = Trim(Request.Form("Title"))
    Keyword = Trim(Request.Form("Keyword"))
    UseLinkUrl = Trim(Request.Form("UseLinkUrl"))
    LinkUrl = Trim(Request.Form("LinkUrl"))
    For i = 1 To Request.Form("Content").Count
        Content = Content & Request.Form("Content")(i)
    Next
    Author = Trim(Request.Form("Author"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
    DefaultPicUrl = Trim(Request.Form("DefaultPicUrl"))
    UploadFiles = Trim(Request.Form("UploadFiles"))
    UpdateTime = PE_CDate(Trim(Request.Form("UpdateTime")))

    'ע����������ֵ
    Status = PE_CLng(Trim(Request.Form("Status")))
    ReceiveUser = ReplaceBadChar(Trim(Request("InceptUser")))
    IncludePic = PE_CLng(Trim(Request.Form("IncludePic")))
    
    Inputer = UserName
    Editor = AdminName


    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then Exit Sub
    
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���ⲻ��Ϊ��</li>"
    Else
        Title = ReplaceText(Title, 2)
    End If
    Keyword = ReplaceBadChar(Keyword)
    If Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������" & ChannelShortName & "�ؼ���</li>"
    Else
        Call SaveKeyword(Keyword)
    End If
    If UseLinkUrl = "Yes" Then
        If LinkUrl = "" Or LCase(LinkUrl) = "http://" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӵ�ַ����Ϊ��</li>"
        Else
            If InStr(LinkUrl, "://") <= 0 And Left(LinkUrl, 1) <> "/" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��վ��ַ���� / ��ͷ��</li>"
            End If
        End If
    Else
        If Content = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���ݲ���Ϊ��</li>"
        End If
    End If
    
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
    Do While Not rsField.EOF
        If rsField("EnableNull") = False Then
            If Trim(Request(rsField("FieldName"))) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������" & rsField("Title") & "��</li>"
            End If
        End If
        rsField.MoveNext
    Loop
    
    If FoundErr = True Then
        Exit Sub
    End If
    
    '����ͼƬJS��ǩ����
    Dim strTemp, strTemp2
    regEx.Pattern = "\<IMG(.[^\<]*)\$\>"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        regEx.Pattern = "\#(.*)\#"
        Set strTemp = regEx.Execute(Match.value)

        For Each Match2 In strTemp
            strTemp2 = Replace(Match2.value, "&amp;", "&")
            strTemp2 = Replace(strTemp2, "#", "")
            strTemp2 = Replace(strTemp2, "&13;&10;", vbCrLf)
            strTemp2 = Replace(strTemp2, "&9;", "vbTab")
            strTemp2 = Replace(strTemp2, "[!", "<")
            strTemp2 = Replace(strTemp2, "!]", ">")
            Content = Replace(Content, Match.value, strTemp2)
        Next
    Next

    Title = PE_HTMLEncode(Title)
    Keyword = Replace("|" & Keyword & "|","||","|")

    '�����Ե�ַת��Ϊ��Ե�ַ
    Dim strSiteUrl
    strSiteUrl = Request.ServerVariables("HTTP_REFERER")
    strSiteUrl = LCase(Left(strSiteUrl, InStrRev(strSiteUrl, "/") - 1))
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/")) & ChannelDir & "/"
    Content = Replace(Content, strSiteUrl, "[InstallDir_ChannelDir]")
    Content = Replace(Content, UploadDir, "{$UploadDir}")

    If Trim(Request.Form("SaveRemotePic")) = "Yes" And EnableSaveRemote = True Then
        Content = ReplaceRemoteUrl(Content)
    End If

    '������δ�����Ϊ�˽������������������µ�Ƶ��Ŀ¼���滻���⡣����˵��ֻ�滻��Ƶ��Ŀ¼��ͷ�ĵ�ַ��������ⲿ��ַ�к���Ƶ��Ŀ¼���Ͳ��滻
    '<a href="/aaa/999.rar">
    '<a href="http://www.baidu.com/aaa/999.rar">
    '<img src="/aaa/999.rar">
    '<img src=/aaa/999.rar>
    '<img src='/aaa/999.rar>

    strSiteUrl = InstallDir & ChannelDir & "/"
    Content = Replace(Content, "'" & strSiteUrl, "'" & "[InstallDir_ChannelDir]")
    Content = Replace(Content, """" & strSiteUrl, """" & "[InstallDir_ChannelDir]")
    Content = Replace(Content, "=" & strSiteUrl, "=" & "[InstallDir_ChannelDir]")

    
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "����")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "��վԭ��")
        
    If Status = 1 Then
        Status = MyStatus
    End If

    Call GetUser(UserName)

    Call SaveVote

    If SpecialID = "" Then
        arrSpecialID = Split("0", ",")
    Else
        arrSpecialID = Split(SpecialID, ",")
    End If
    PresentExp = Int(PresentExp * PresentExpTimes)

    Set rsArticle = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
        If Session("Title") = Title And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>�벻Ҫ�ظ����ͬһ" & ChannelItemUnit & ChannelShortName & "</li>"
            Exit Sub
        Else
            Session("Title") = Title
            Session("AddTime") = Now()
            ArticleID = GetNewID("PE_Article", "ArticleID")
            
            For i = 0 To UBound(arrSpecialID)
                Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & ArticleID & "," & PE_CLng(arrSpecialID(i)) & ")")
            Next
            sql = "select top 1 * from PE_Article"
            rsArticle.Open sql, Conn, 1, 3
            rsArticle.addnew
            rsArticle("ArticleID") = ArticleID
            rsArticle("ChannelID") = ChannelID
            rsArticle("Inputer") = Inputer

            If UserID <> "" And UserID > 0 Then
                Dim blogid
                Set blogid = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If blogid.BOF And blogid.EOF Then
                    rsArticle("BlogID") = 0
                Else
                    rsArticle("BlogID") = blogid("ID")
                End If
                Set blogid = Nothing
            End If
            
            If ReceiveUser <> "" And Status = 3 Then
                rsArticle("Receive") = True
                Call Add_User_UnsignedItems(ArticleID, ReceiveUser)
            Else
                rsArticle("Receive") = False
            End If

            If Status = 3 Then
                Conn.Execute ("update PE_User set UserExp=UserExp+" & PresentExp & " where UserName='" & Inputer & "'")
            End If
        End If
        
    ElseIf Action = "SaveModify" Then
        If ArticleID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷ��ArticleID��ֵ</li>"
        Else
            ArticleID = PE_CLng(ArticleID)
            sql = "select * from PE_Article where ArticleID=" & ArticleID
            rsArticle.Open sql, Conn, 1, 3
            If rsArticle.BOF And rsArticle.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Ҳ�����" & ChannelShortName & "�������Ѿ���������ɾ����</li>"
            Else
            
                'ɾ�����ɵ��ļ�����Ϊ���ɵ��ļ����ܻ����Ÿ���ʱ�䣬����Ȩ�޵ȷ����仯������������ļ�
                If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
                    Dim tClass, ArticlePath
                    Set tClass = Conn.Execute("select ParentDir,ClassDir from PE_Class where ClassID=" & rsArticle("ClassID") & "")
                    If tClass.BOF And tClass.EOF Then
                        ParentDir = "/"
                        ClassDir = ""
                    Else
                        ParentDir = tClass("ParentDir")
                        ClassDir = tClass("ClassDir")
                    End If
                    ArticlePath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsArticle("UpdateTime"), rsArticle("ArticleID"))
                    If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
                        DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
                    End If
                    If rsArticle("PaginationType") > 0 Then
                        DelSerialFiles Server.MapPath(ArticlePath) & "_*.*"
                    End If
                End If
                If rsArticle("Inputer") <> UserName And rsArticle("Status") <> Status And (Status = -2 Or Status = 3) Then
                    Call SendEmailOfCheck(rsArticle("Inputer"), rsArticle)
                End If

                Call UpdateUserData(0, rsArticle("Inputer"), 0, 0)
            
                If rsArticle("Status") < 3 And Status = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp+" & rsArticle("PresentExp") & " where UserName='" & rsArticle("Inputer") & "'")
                End If
                If rsArticle("Status") = 3 And Status < 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp-" & rsArticle("PresentExp") & " where UserName='" & rsArticle("Inputer") & "'")
                End If
    
                Dim rsInfo, sqlInfo, j
                i = 0
                sqlInfo = "select * from PE_InfoS where ModuleType=" & ModuleType & " and ItemID=" & ArticleID & " order by InfoID"
                Set rsInfo = Server.CreateObject("adodb.recordset")
                rsInfo.Open sqlInfo, Conn, 1, 3
                Do While Not rsInfo.EOF
                    If i > UBound(arrSpecialID) Then
                        rsInfo.Delete
                    Else
                        rsInfo("SpecialID") = arrSpecialID(i)
                    End If
                    rsInfo.Update
                    rsInfo.MoveNext
                    i = i + 1
                Loop
                rsInfo.Close
                Set rsInfo = Nothing
                If (i - 1) < UBound(arrSpecialID) Then
                    For j = i To UBound(arrSpecialID)
                        If PE_CLng(arrSpecialID(j)) <> 0 Then
                            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & ArticleID & "," & PE_CLng(arrSpecialID(j)) & ")")
                        End If
                    Next
                End If
                
                If ReceiveUser = "" Or Status <> 3 Then
                    rsArticle("Receive") = False
                    If rsArticle("ReceiveUser") <> "" Then
                        Call Del_User_UnsignedItems(ArticleID, rsArticle("ReceiveUser"))
                    End If
                Else
                    rsArticle("Receive") = True
                    If rsArticle("ReceiveUser") <> "" Then
                        Call Del_User_UnsignedItems(ArticleID, rsArticle("ReceiveUser"))
                    End If
                    Call Add_User_UnsignedItems(ArticleID, ReceiveUser)
                End If
                        
            End If
        End If
    End If

    rsArticle("ClassID") = ClassID
    rsArticle("Title") = Title
    rsArticle("TitleIntact") = Trim(Request.Form("TitleIntact"))
    rsArticle("Subheading") = Trim(Request.Form("Subheading"))
    rsArticle("TitleFontColor") = Trim(Request.Form("TitleFontColor"))
    rsArticle("TitleFontType") = PE_CLng(Trim(Request.Form("TitleFontType")))
    rsArticle("Intro") = Trim(Request.Form("Intro"))
    rsArticle("Content") = Content
    rsArticle("Keyword") = Keyword
    rsArticle("Author") = Author
    rsArticle("CopyFrom") = CopyFrom
    rsArticle("LinkUrl") = LinkUrl
    rsArticle("Editor") = Editor
    rsArticle("IncludePic") = IncludePic
    rsArticle("ShowCommentLink") = PE_CBool(Trim(Request.Form("ShowCommentLink")))
    rsArticle("Status") = Status
    rsArticle("OnTop") = PE_CBool(Trim(Request.Form("OnTop")))
    rsArticle("Elite") = PE_CBool(Trim(Request.Form("Elite")))
    If Action = "SaveModifyAsAdd" Then 
        rsArticle("Hits") = 0
	Else 
	    rsArticle("Hits") = PE_CLng(Trim(Request.Form("Hits")))
    End IF	
    rsArticle("Stars") = PE_CLng(Trim(Request.Form("Stars")))
    rsArticle("UpdateTime") = UpdateTime
    rsArticle("CreateTime") = UpdateTime
    rsArticle("PaginationType") = PE_CLng(Trim(Request.Form("PaginationType")))
    rsArticle("MaxCharPerPage") = PE_CLng(Trim(Request.Form("MaxCharPerPage")))
    rsArticle("SkinID") = PE_CLng(Trim(Request.Form("SkinID")))
    rsArticle("TemplateID") = PE_CLng(Trim(Request.Form("TemplateID")))
    rsArticle("DefaultPicUrl") = DefaultPicUrl
    rsArticle("UploadFiles") = UploadFiles
    rsArticle("Deleted") = False
    rsArticle("PresentExp") = PresentExp

    rsArticle("Copymoney") = PE_CDbl(Trim(Request.Form("CopyMoney"))) '���
    rsArticle("IsPayed") = False
    rsArticle("Beneficiary") = Trim(Request.Form("Beneficiary"))    '��������� ���������֮���á���������

    rsArticle("InfoPurview") = PE_CLng(Trim(Request.Form("InfoPurview")))
    rsArticle("arrGroupID") = ReplaceBadChar(Trim(Request.Form("GroupID")))
    rsArticle("InfoPoint") = PE_CLng(Trim(Request.Form("InfoPoint")))
    rsArticle("ChargeType") = PE_CLng(Trim(Request.Form("ChargeType")))
    rsArticle("PitchTime") = PE_CLng(Trim(Request.Form("PitchTime")))
    rsArticle("ReadTimes") = PE_CLng(Trim(Request.Form("ReadTimes")))
    rsArticle("DividePercent") = PE_CLng(Trim(Request.Form("DividePercent")))

    rsArticle("ReceiveUser") = ReceiveUser
    rsArticle("Received") = ""
    rsArticle("AutoReceiveTime") = PE_CLng(Request("AutoReceiveTime"))
    rsArticle("ReceiveType") = PE_CLng(Request("ReceiveType")) '���ǩ�����µ����ͣ�0 Ϊ���У�1Ϊ˽��
    rsArticle("VoteID") = VoteID

    If Not (rsField.BOF And rsField.EOF) Then
        rsField.MoveFirst
        Do While Not rsField.EOF
            If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                rsArticle(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
            End If
            rsField.MoveNext
        Loop
    End If
    Set rsField = Nothing

    rsArticle.Update
    rsArticle.Close
    Set rsArticle = Nothing
    Call UpdateChannelData(ChannelID)
    If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
        Call UpdateUserData(0, Inputer, 0, 0)
    End If
    Response.Write "<br><br>"
    Response.Write "<table class='border' align='center' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "  <tr class='title'> "
    Response.Write "    <td  height='22' align='center' colspan='2'> "
    If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
        Response.Write "<b>���" & ChannelShortName & "�ɹ�</b>"
    Else
        Response.Write "<b>�޸�" & ChannelShortName & "�ɹ�</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>������Ŀ��</strong></td>"
    Response.Write "          <td width='400'>" & ShowClassPath() & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "���⣺</strong></td>"
    Response.Write "          <td width='400'>" & Title & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�</strong></td>"
    Response.Write "          <td width='400'>" & Author & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Դ��</strong></td>"
    Response.Write "          <td width='400'>" & CopyFrom & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>�� �� �֣�</strong></td>"
    Response.Write "          <td width='400'>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "״̬��</strong></td>"
    Response.Write "          <td width='400'>"
    If Status = -1 Then
        Response.Write "�ݸ�"
    ElseIf Status = -2 Then
        Response.Write "�˸�"
    Else
        Response.Write arrStatus(Status)
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg' align='center'>"
    Response.Write "    <td height='30' colspan='2'>"
    Response.Write "��<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Modify&ArticleID=" & ArticleID & "'>�޸ı���</a>��&nbsp;"
    Response.Write "��<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>�������" & ChannelShortName & "</a>��&nbsp;"
    Response.Write "��<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "����</a>��&nbsp;"
    Response.Write "��<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & "'>�鿴" & ChannelShortName & "����</a>��"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf

    Session("Keyword") = Trim(Request("Keyword"))
    Session("Author") = Author
    Session("CopyFrom") = CopyFrom
    Session("PaginationType") = PE_CLng(Trim(Request("PaginationType")))
    Session("SkinID") = PE_CLng(Trim(Request("SkinID")))
    Session("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
    
    Call ClearSiteCache(0)
    Call CreateAllJS

    If Status = 3 And PE_CLng(Trim(Request("ReceiveType"))) = 0 And UseCreateHTML > 0 And ObjInstalled_FSO = True And Trim(Request.Form("CreateImmediate")) = "Yes" Then
        Response.Write "<br><iframe id='CreateArticle' width='100%' height='210' frameborder='0' src='Admin_CreateArticle.asp?ChannelID=" & ChannelID & "&Action=CreateArticle2&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&ArticleID=" & ArticleID & "&ShowBack=No'></iframe>"
    End If
End Sub

Sub Show()
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��" & ChannelShortName & "ID��</li>"
        Exit Sub
    End If
    
    Dim rsArticle, PurviewChecked, PurviewChecked2
    PurviewChecked = False
    PurviewChecked2 = False
    Set rsArticle = Conn.Execute("select * from PE_Article where ArticleID=" & ArticleID & "")
    If rsArticle.BOF And rsArticle.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "��</li>"
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If
    ClassID = rsArticle("ClassID")

    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If

    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, ArticleID)

    Call WriteJS_Show

    Response.Write "<br>�����ڵ�λ�ã�&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "����</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Conn.Execute(sqlPath)
        Do While Not rsPath.EOF
            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;�鿴" & ChannelShortName & "���ݣ�"
    Select Case rsArticle("IncludePic")
        Case 1
            Response.Write "<font color=blue>" & ArticlePro1 & "</font>"
        Case 2
            Response.Write "<font color=blue>" & ArticlePro2 & "</font>"
        Case 3
            Response.Write "<font color=blue>" & ArticlePro3 & "</font>"
        Case 4
            Response.Write "<font color=blue>" & ArticlePro4 & "</font>"
    End Select
    
    Response.Write rsArticle("Title") & "<br><br>"



    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>����ר��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�շ���Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>ǩ����Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>�����Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    
    Response.Write "  <tr align='center' class='tdbg'>"
    Response.Write "    <td height='40' colspan='2'>"
    If Trim(rsArticle("TitleIntact")) <> "" Then
        Response.Write "<font size='4'><b>" & rsArticle("TitleIntact") & "</b></font>"
    Else
        Response.Write "<font size='4'><b>" & rsArticle("Title") & "</b></font>"
    End If
    If Trim(rsArticle("Subheading")) <> "" Then
        Response.Write "<br>" & rsArticle("Subheading")
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center' class='tdbg'>"
    Response.Write "    <td colspan='2'>"
    Dim Author, CopyFrom
    Author = rsArticle("Author")
    CopyFrom = rsArticle("CopyFrom")
    Response.Write "���ߣ�" & Author & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "��Դ��" & CopyFrom
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�������" & rsArticle("Hits") & "&nbsp;&nbsp;&nbsp;&nbsp;����ʱ�䣺" & FormatDateTime(rsArticle("UpdateTime"), 2) & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "���ԣ�"
    If rsArticle("OnTop") = True Then
        Response.Write "<font color=blue>��</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Hits") >= HitsOfHot Then
        Response.Write "<font color=red>��</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Elite") = True Then
        Response.Write "<font color=green>��</font>"
    Else
        Response.Write "&nbsp;&nbsp;"
    End If
    Response.Write "&nbsp;&nbsp;<font color='#009900'>" & String(rsArticle("Stars"), "��") & "</font>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan='2'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='5'>"
    Response.Write "        <tr>"
    Response.Write "          <td height='200' valign='top'>"
    If Trim(rsArticle("LinkUrl")) <> "" Then
        Response.Write "<p align='center'><br><br><br><font color=red>��" & ChannelShortName & "�������ⲿ" & ChannelShortName & "���ݡ����ӵ�ַΪ��<a href='" & rsArticle("LinkUrl") & "' target='_blank'>" & rsArticle("LinkUrl") & "</a></font></p>"
    Else
        Response.Write "<p>" & Replace(Replace(FilterBadTag(rsArticle("Content"), rsArticle("Inputer")), "[InstallDir_ChannelDir]", InstallDir & ChannelDir & "/"), "{$UploadDir}", UploadDir) & "</p>"
    End If
    Response.Write "       </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr  align='right' class='tdbg'>"
    Response.Write "    <td colspan='2'>"
    Response.Write ChannelShortName & "¼�룺<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Field=Inputer&Keyword=" & rsArticle("Inputer") & "'>" & rsArticle("Inputer") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;���α༭��"
    If rsArticle("Status") = 3 Then
        Response.Write rsArticle("Editor")
    Else
        Response.Write "��"
    End If
    Response.Write " </td>"
    Response.Write "  </tr>"
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Special(arrSpecialID, " disabled")

    Call ShowTabs_Purview_Modify("�Ķ�", rsArticle, " disabled")
    
    
    Dim NotReceiveUser, arrUser, i
    If rsArticle("Receive") = True Then
        If rsArticle("Received") = "" Then
            NotReceiveUser = rsArticle("ReceiveUser")
        Else
            NotReceiveUser = ""
            arrUser = Split(rsArticle("ReceiveUser"), ",")
            For i = 0 To UBound(arrUser)
                If FoundInArr(rsArticle("Received"), arrUser(i), "|") = False Then
                    If NotReceiveUser = "" Then
                        NotReceiveUser = arrUser(i)
                    Else
                        NotReceiveUser = NotReceiveUser & "," & arrUser(i)
                    End If
                End If
            Next
        End If
    End If
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>Ҫ��ǩ���û���</td>"
    Response.Write "            <td style='width:600; word-wrap:break-word;'>" & rsArticle("ReceiveUser") & "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�Ѿ�ǩ���û���</td>"
    Response.Write "            <td style='width:600; word-wrap:break-word;'>" & rsArticle("Received") & "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��δǩ���û���</td>"
    Response.Write "            <td style='width:600; word-wrap:break-word;'>" & NotReceiveUser & "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>ǩ�շ�ʽ��</td>"
    Response.Write "            <td><select name='AutoReceiveTime' disabled>"
    Response.Write "                      <option value='0'"
    If rsArticle("AutoReceiveTime") = "0" Then Response.Write " selected"
    Response.Write ">�ֶ�ǩ��</option>"
    Response.Write "                      <option value='5'"
    If rsArticle("AutoReceiveTime") = "5" Then Response.Write " selected"
    Response.Write ">5���Ӻ�</option>"
    Response.Write "                      <option value='10'"
    If rsArticle("AutoReceiveTime") = "10" Then Response.Write " selected"
    Response.Write ">10���Ӻ�</option>"
    Response.Write "                      <option value='30'"
    If rsArticle("AutoReceiveTime") = "30" Then Response.Write " selected"
    Response.Write ">30���Ӻ�</option>"
    Response.Write "                      <option value='60'"
    If rsArticle("AutoReceiveTime") = "60" Then Response.Write " selected"
    Response.Write ">1���Ӻ�</option>"
    Response.Write "                      <option value='120'"
    If rsArticle("AutoReceiveTime") = "120" Then Response.Write " selected"
    Response.Write ">2���Ӻ�</option>"
    Response.Write "                      <option value='300'"
    If rsArticle("AutoReceiveTime") = "300" Then Response.Write " selected"
    Response.Write ">5���Ӻ�</option>"
    Response.Write "                    </select>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ĵ����ͣ�</td>"
    Response.Write "            <td><select name='ReceiveType' id='ReceiveType' disabled>"
    Response.Write "                      <option value='0'"
    If rsArticle("ReceiveType") = "0" Then Response.Write " selected"
    Response.Write ">�����ĵ�</option>"
    Response.Write "                      <OPTION value='1'"
    If rsArticle("ReceiveType") = "1" Then Response.Write " selected"
    Response.Write ">ר���ĵ�</OPTION>"
    Response.Write "                    </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    
    Dim WordNum, strArticle, strSql, rsChannel, payNum
    '���˿ո�HTML�ַ�����������
    WordNum = getWordNumber(rsArticle("Content"))
    
    If MoneyPerKw <= 0 Then
       payNum = 0
    Else
       payNum = MoneyPerKw / 1000 * WordNum
    End If
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>����������</td><td>" & WordNum & " ��</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>֧����׼��</td><td>" & MoneyPerKw & " Ԫ/ǧ��</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>���Ƹ�ѣ�</td><td>" & payNum & " Ԫ</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>ȷ��֧����ѣ�</td><td>" & rsArticle("CopyMoney") & " Ԫ</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��������ߣ�</td><td>" & rsArticle("Inputer") & "</td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    

    Call ShowTabs_MyField_View(rsArticle)

    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf



    Response.Write "<form name='formA' method='get' action='Admin_Article.asp'><p align='center'>"
    Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='hidden' name='ArticleID' value='" & ArticleID & "'>"
    Response.Write "<input type='hidden' name='Action' value=''>" & vbCrLf

    If rsArticle("Deleted") = False Then
        PurviewChecked = CheckClassPurview("Manage", ClassID)
        PurviewChecked2 = CheckClassPurview("Check", ClassID)
        If (rsArticle("Inputer") = UserName And rsArticle("Status") = 0) Or PurviewChecked = True Then
            Response.Write "<input type='submit' name='submit' value='�޸�/���' onclick=""document.formA.Action.value='Modify'"">&nbsp;&nbsp;"
            Response.Write "<input type='submit' name='submit' value=' ɾ �� ' onclick=""document.formA.Action.value='Del'"">&nbsp;&nbsp;"
        End If
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input type='submit' name='submit' value=' �� �� ' onclick=""document.formA.Action.value='MoveToClass'"">&nbsp;&nbsp;"
        End If
        If PurviewChecked2 = True Then
            If rsArticle("Status") > -1 Then
                Response.Write "<input type='submit' name='submit' value='ֱ���˸�' onclick=""document.formA.Action.value='Reject'"">&nbsp;&nbsp;"
            End If
            If rsArticle("Status") < MyStatus Then
                Response.Write "<input type='submit' name='submit' value='" & arrStatus(MyStatus) & "' onclick=""document.formA.Action.value='SetPassed'"">&nbsp;&nbsp;"
            End If
            If rsArticle("Status") >= MyStatus Then
                Response.Write "<input type='submit' name='submit' value='ȡ�����' onclick=""document.formA.Action.value='CancelPassed'"">&nbsp;&nbsp;"
            End If
        End If
        If PurviewChecked = True Then
            If rsArticle("OnTop") = False Then
                Response.Write "<input type='submit' name='submit' value='��Ϊ�̶�' onclick=""document.formA.Action.value='SetOnTop'"">&nbsp;&nbsp;"
            Else
                Response.Write "<input type='submit' name='submit' value='ȡ���̶�' onclick=""document.formA.Action.value='CancelOnTop'"">&nbsp;&nbsp;"
            End If
            If rsArticle("Elite") = False Then
                Response.Write "<input type='submit' name='submit' value='��Ϊ�Ƽ�' onclick=""document.formA.Action.value='SetElite'"">"
            Else
                Response.Write "<input type='submit' name='submit' value='ȡ���Ƽ�' onclick=""document.formA.Action.value='CancelElite'"">"
            End If
        End If
    Else
        If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
            Response.Write "<input type='submit' name='submit' value='����ɾ��' onclick=""if(confirm('ȷ��Ҫ����ɾ����" & ChannelShortName & "�𣿳���ɾ�����޷���ԭ��')==true){document.formA.Action.value='ConfirmDel';}"">&nbsp;&nbsp;"
            Response.Write "<input type='submit' name='submit' value=' �� ԭ ' onclick=""document.formA.Action.value='Restore'"">"
        End If
    End If
    Response.Write "</p></form>"

    rsArticle.Close
    Set rsArticle = Nothing

    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='0'><tr class='tdbg'><td>"
    Response.Write "<li>��һ" & ChannelItemUnit & ChannelShortName & "��"
    Dim rsPrev
    Set rsPrev = Conn.Execute("Select Top 1 A.ArticleID,A.Title,C.ClassID,C.ClassName from PE_Article A left join PE_Class C On A.ClassID=C.ClassID Where A.ChannelID=" & ChannelID & " and A.Deleted=" & PE_False & " and A.ArticleID<" & ArticleID & " order by A.ArticleID desc")
    If rsPrev.EOF Then
        Response.Write "û����"
    Else
        Response.Write "[<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPrev("ClassID") & "'>" & rsPrev("ClassName") & "</a>] <a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsPrev("ArticleID") & "'>" & rsPrev("Title") & "</a>"
    End If
    rsPrev.Close
    Set rsPrev = Nothing
    Response.Write "</li></td</tr><tr class='tdbg'><td><li>��һ" & ChannelItemUnit & ChannelShortName & "��"
    Dim rsNext
    Set rsNext = Conn.Execute("Select Top 1 A.ArticleID,A.Title,C.ClassID,C.ClassName from PE_Article A left join PE_Class C On A.ClassID=C.ClassID Where A.ChannelID=" & ChannelID & " and A.Deleted=" & PE_False & " and A.ArticleID>" & ArticleID & " order by A.ArticleID asc")
    If rsNext.EOF Then
        Response.Write "û����"
    Else
        Response.Write "[<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsNext("ClassID") & "'>" & rsNext("ClassName") & "</a>] <a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsNext("ArticleID") & "'>" & rsNext("Title") & "</a>"
    End If
    rsNext.Close
    Set rsNext = Nothing
    Response.Write "</li></td></tr></table><br>" & vbCrLf

    Dim InfoType
    InfoType = PE_CLng(Trim(Request("InfoType")))

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'><td"
    If InfoType = 0 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_Article.asp?Action=Show&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&InfoType=0'"""
    End If
    Response.Write ">�������</td><td"
    If InfoType = 1 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_Article.asp?Action=Show&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&InfoType=1'"""
    End If
    Response.Write ">����շ�</td>"
    Response.Write "<td>&nbsp;</td></tr></table>"
    
    strFileName = "Admin_Article.asp?Action=Show&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&InfoType=" & InfoType
    
    Select Case InfoType
    Case 0
        Call ShowComment(ArticleID)
    Case 1
        Call ShowConsumeLog(ArticleID)
    End Select
End Sub

Sub WriteJS_Show()
    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "function resizepic(thispic){" & vbCrLf
    Response.Write "  if(thispic.width>600) thispic.width=600;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function bbimg(o){" & vbCrLf
    Response.Write "  var zoom=parseInt(o.style.zoom, 10)||100;" & vbCrLf
    Response.Write "  zoom+=event.wheelDelta/12;" & vbCrLf
    Response.Write "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    Response.Write "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    Response.Write "  return false;" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

Sub Preview()
    Call WriteJS_Show
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td width='400' height='22'>"

    If ClassID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��������Ŀ</li>"
        Exit Sub
    ElseIf ClassID > 0 Then
        Set tClass = Conn.Execute("select ClassName,ParentID,ParentPath from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ</li>"
            Set tClass = Nothing
            Exit Sub
        Else
            ClassName = tClass(0)
            ParentID = tClass(1)
            ParentPath = tClass(2)
        End If
        Set tClass = Nothing
        If ParentID > 0 Then
            Dim sqlPath, rsPath
            sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
            Set rsPath = Conn.Execute(sqlPath)
            Do While Not rsPath.EOF
                Response.Write rsPath(1) & "&nbsp;&gt;&gt;&nbsp;"
                rsPath.MoveNext
            Loop
            rsPath.Close
            Set rsPath = Nothing
        End If
        Response.Write ClassName & "&nbsp;&gt;&gt;&nbsp;"
    End If

    Select Case Trim(Request("IncludePic"))
        Case 1
            Response.Write "<font color=blue>" & ArticlePro1 & "</font>"
        Case 2
            Response.Write "<font color=blue>" & ArticlePro2 & "</font>"
        Case 3
            Response.Write "<font color=blue>" & ArticlePro3 & "</font>"
        Case 4
            Response.Write "<font color=blue>" & ArticlePro4 & "</font>"
    End Select
    
    Response.Write PE_HTMLEncode(Request("Title"))
    Response.Write " </td>"
    Response.Write "    <td width='50' height='22' align='right'>"
    If LCase(Request("OnTop")) = "yes" Then
        Response.Write "��&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If LCase(Request("Hot")) = "yes" Then
        Response.Write "��&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If LCase(Request("Elite")) = "yes" Then
        Response.Write "��"
    Else
        Response.Write "&nbsp;&nbsp;"
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3' align='center'><font size='4'>"
    If Trim(Request("TitleIntact")) <> "" Then
        Response.Write "<font size='4'><b>" & PE_HTMLEncode(Request("TitleIntact")) & "</b></font>"
    Else
        Response.Write "<font size='4'><b>" & PE_HTMLEncode(Request("Title")) & "</b></font>"
    End If
    If Trim(Request("Subheading")) <> "" Then
        Response.Write "<br>" & PE_HTMLEncode(Request("Subheading"))
    End If

    Response.Write "</font></td></tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3' align='center'>���ߣ�" & PE_HTMLEncode(Request("Author")) & "&nbsp;&nbsp;&nbsp;&nbsp;ת���ԣ�" & PE_HTMLEncode(Request("CopyFrom")) & "&nbsp;&nbsp;&nbsp;&nbsp;�������0&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "¼�룺" & UserName & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3'><p>" & Request("Content") & "</p></td></tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>��<a href='javascript:window.close();'>�رմ���</a>��</p>"
End Sub


'******************************************************************************************
'����Ϊ������������ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub Batch()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If

    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    
    ArticleID = Replace(ArticleID, " ", "")
    Response.Write "<form method='POST' name='myform' action='Admin_Article.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><b>�����޸�" & ChannelShortName & "����</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center'>"
    Response.Write "      <td class='tdbg' valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td align='center'><b>" & ChannelShortName & "��Χ</b></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='BatchType' value='1' checked>ָ��" & ChannelShortName & "ID��<br>"
    Response.Write "              <input type='text' name='BatchArticleID' value='" & ArticleID & "' size='28'><br>"
    Response.Write "              <input type='radio' name='BatchType' value='2'>ָ����Ŀ��" & ChannelShortName & "��<br>"
    Response.Write "              <select name='BatchClassID' size='2' multiple style='height:280px;width:180px;'>" & GetClass_Option(0, 0) & "</select><br><div align='center'>"
    Response.Write "      <input type='button' name='Submit' value='  ѡ��������Ŀ  ' onclick='SelectAll()'><br>"
    Response.Write "      <input type='button' name='Submit' value='ȡ��ѡ��������Ŀ' onclick='UnSelectAll()'></div></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "      <td valign='top' align='left'><br>"
    
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyCommentLink' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���⣺</td>"
    Response.Write "            <td><input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'> �б���ʾʱ�ڱ�������ʾ�����ۡ�����"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyAuthor' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���ߣ�</td>"
    Response.Write "            <td><input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='15' maxlength='30'> " & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyCopyFrom' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "��Դ��</td>"
    Response.Write "            <td><input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='15' maxlength='50'> " & GetCopyFromList("Admin", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyPaginationType' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>���ݷ�ҳ��ʽ��</td>"
    Response.Write "            <td><select name='PaginationType' id='PaginationType'>"
    Response.Write "                <option value='0' selected>����ҳ</option>"
    Response.Write "                <option value='1'>�Զ���ҳ</option>"
    Response.Write "                <option value='2'>�ֶ���ҳ</option>"
    Response.Write "              </select>"
    Response.Write "              �Զ���ҳʱ��ÿҳ��Լ�ַ���������HTML����ұ������100����<input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value='10000' size='8' maxlength='8'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Call ShowBatchCommon
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Purview_Batch("�Ķ�")
    Call ShowTabs_MyField_Batch

    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <br><b>˵����</b><br>1����Ҫ�����޸�ĳ�����Ե�ֵ������ѡ�������ĸ�ѡ��Ȼ�����趨����ֵ��<br>2��������ʾ������ֵ����ϵͳĬ��ֵ������ѡ" & ChannelShortName & "�����������޹�<br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='DoBatch'>"
    Response.Write "    <input name='add' type='submit'  id='Add' value=' ִ�������� ' style='cursor:hand;'>&nbsp; "
    Response.Write "    <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p>"
    Response.Write "  <br>"
    Response.Write "</form>"
End Sub

Sub DoBatch()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim rs, sql, BatchType, BatchArticleID, BatchClassID, rsField
    Dim Author, ShowCommentLink, CopyFrom, PaginationType, MaxCharPerPage
    Dim Keyword, OnTop, Elite, Stars, Hits, UpdateTime, SkinID, TemplateID
    Dim InfoPurview, arrGroupID, InfoPoint, ChargeType, PitchTime, ReadTimes, DividePercent
    
    BatchType = PE_CLng(Trim(Request("BatchType")))
    BatchArticleID = Trim(Request.Form("BatchArticleID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    Author = Trim(Request.Form("Author"))
    ShowCommentLink = Trim(Request.Form("ShowCommentLink"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
    PaginationType = PE_CLng(Trim(Request.Form("PaginationType")))
    MaxCharPerPage = PE_CLng(Trim(Request.Form("MaxCharPerPage")))
    Keyword = Trim(Request.Form("Keyword"))
    OnTop = Trim(Request.Form("OnTop"))
    Elite = Trim(Request.Form("Elite"))
    Stars = PE_CLng(Trim(Request.Form("Stars")))
    Hits = PE_CLng(Trim(Request.Form("Hits")))
    UpdateTime = PE_CDate(Trim(Request.Form("UpdateTime")))
    SkinID = PE_CLng(Trim(Request.Form("SkinID")))
    TemplateID = PE_CLng(Trim(Request.Form("TemplateID")))
    
    InfoPurview = PE_CLng(Trim(Request.Form("InfoPurview")))
    arrGroupID = ReplaceBadChar(Trim(Request.Form("GroupID")))
    InfoPoint = PE_CLng(Trim(Request.Form("InfoPoint")))
    ChargeType = PE_CLng(Trim(Request.Form("ChargeType")))
    PitchTime = PE_CLng(Trim(Request.Form("PitchTime")))
    ReadTimes = PE_CLng(Trim(Request.Form("ReadTimes")))
    DividePercent = PE_CLng(Trim(Request.Form("DividePercent")))

    If BatchType = 1 Then
        If IsValidID(BatchArticleID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����޸ĵ�" & ChannelShortName & "��ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����޸ĵ�" & ChannelShortName & "����Ŀ</li>"
        End If
    End If
    If Trim(Request("ModifyKeyword")) = "Yes" And Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������" & ChannelShortName & "�ؼ���</li>"
    End If
    If Trim(Request("ModifyPaginationType")) = "Yes" And PaginationType = 1 And MaxCharPerPage = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ���Զ���ҳʱ��ÿҳ��Լ�ַ������������0</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "����")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "��վԭ��")

    If ShowCommentLink = "Yes" Then
        ShowCommentLink = True
    Else
        ShowCommentLink = False
    End If
    Keyword = "|" & ReplaceBadChar(Keyword) & "|"
    If OnTop = "Yes" Then
        OnTop = True
    Else
        OnTop = False
    End If
    If Elite = "Yes" Then
        Elite = True
    Else
        Elite = False
    End If

    Set rs = Server.CreateObject("adodb.recordset")
    If BatchType = 1 Then
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ArticleID in (" & BatchArticleID & ")"
    Else
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ClassID in (" & BatchClassID & ")"
    End If
    rs.Open sql, Conn, 1, 3
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
    Do While Not rs.EOF
        If Trim(Request("ModifyAuthor")) = "Yes" Then rs("Author") = Author
        If Trim(Request("ModifyCopyFrom")) = "Yes" Then rs("CopyFrom") = CopyFrom
        If Trim(Request("ModifyCommentLink")) = "Yes" Then rs("ShowCommentLink") = ShowCommentLink
        If Trim(Request("ModifyPaginationType")) = "Yes" Then
            rs("PaginationType") = PaginationType
            rs("MaxCharPerPage") = MaxCharPerPage
        End If
        If Trim(Request("ModifyKeyword")) = "Yes" Then rs("Keyword") = Keyword
        If Trim(Request("ModifyOnTop")) = "Yes" Then rs("OnTop") = OnTop
        If Trim(Request("ModifyElite")) = "Yes" Then rs("Elite") = Elite
        If Trim(Request("ModifyStars")) = "Yes" Then rs("Stars") = Stars
        If Trim(Request("ModifyHits")) = "Yes" Then rs("Hits") = Hits
        If Trim(Request("ModifyUpdateTime")) = "Yes" Then rs("UpdateTime") = DateAdd("d", DateDiff("d", rs("UpdateTime"), UpdateTime), rs("UpdateTime"))
        If Trim(Request("ModifySkin")) = "Yes" Then rs("SkinID") = SkinID
        If Trim(Request("ModifyTemplate")) = "Yes" Then rs("TemplateID") = TemplateID

        If Trim(Request("ModifyInfoPurview")) = "Yes" Then
            rs("InfoPurview") = InfoPurview
            rs("arrGroupID") = arrGroupID
        End If
        If Trim(Request("ModifyInfoPoint")) = "Yes" Then rs("InfoPoint") = InfoPoint
        If Trim(Request("ModifyChargeType")) = "Yes" Then
            rs("ChargeType") = ChargeType
            rs("PitchTime") = PitchTime
            rs("ReadTimes") = ReadTimes
        End If
        If Trim(Request("ModifyDividePercent")) = "Yes" Then rs("DividePercent") = DividePercent

        If Not (rsField.BOF And rsField.EOF) Then
            rsField.MoveFirst
            Do While Not rsField.EOF
                If Trim(Request("Modify" & Trim(rsField("FieldName")))) = "Yes" Then
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rs(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
                    End If
                End If
                rsField.MoveNext
            Loop
        End If
        rs("CreateTime") = rs("UpdateTime")

        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsField.Close
    Set rsField = Nothing
    Call ClearSiteCache(0)

    Call WriteSuccessMsg("�����޸�" & ChannelShortName & "���Գɹ���", "Admin_Article.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub
'=================================================
'��������BatchReplace
'��  �ã������滻
'=================================================
Sub BatchReplace()

    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If

    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    ArticleID = Replace(ArticleID, " ", "")
    Response.Write "<form method='POST' name='myform' action='Admin_Article.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><b>�����滻" & ChannelShortName & "����</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "       <td class='tdbg' valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td align='center'><b>" & ChannelShortName & "��Χ</b></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='BatchType' value='1' checked>ָ��" & ChannelShortName & "ID��<br>"
    Response.Write "              <input type='text' name='BatchArticleID' value='" & ArticleID & "' size='28'><br>"
    Response.Write "              <input type='radio' name='BatchType' value='2'>ָ����Ŀ��" & ChannelShortName & "��<br>"
    Response.Write "              <select name='BatchClassID' size='2' multiple style='height:280px;width:180px;'>" & GetClass_Option(0, 0) & "</select><br><div align='center'>"
    Response.Write "              <input type='button' name='Submit' value='  ѡ��������Ŀ  ' onclick='SelectAll()'><br>"
    Response.Write "              <input type='button' name='Submit' value='ȡ��ѡ��������Ŀ' onclick='UnSelectAll()'></div></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "     </td>" & vbCrLf
    Response.Write "      <td valign='top'>" & vbCrLf
    Response.Write "       <table width='100%' height='400' border='0' cellpadding='0' cellspacing='1'>"
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>�滻���ݣ�&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='checkbox' NAME='ItemBatchTitle'  value='yes' >" & ChannelShortName & "����&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='checkbox' NAME='ItemBatchContent'  value='yes' checked>" & ChannelShortName & "����</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>�滻���ͣ�&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='ItemBatchType' onClick=""javascript:PE_ItemReplaceStart.style.display='none';PE_ItemReplaceEnd.style.display='none';PE_ItemReplace.style.display='';"" value='1' checked>���滻&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='ItemBatchType' onClick=""javascript:PE_ItemReplaceStart.style.display='';PE_ItemReplaceEnd.style.display='';PE_ItemReplace.style.display='none';"" value='2' >�߼��滻</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplace' style='display:'> " & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5' align=""right""><strong> Ҫ�滻���ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplace' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplaceStart' style='display:none'> " & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5' align=""right"" ><strong> Ҫ�滻�Ŀ�ʼ�ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplaceStart' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplaceEnd' style='display:none'> " & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5' align=""right"" ><strong> Ҫ�滻�Ľ����ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplaceEnd' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplaceResult' style='display:'>" & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5'  align=""right""><strong> Ҫ�滻����ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplaceResult' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>�Ƿ�����ǰ׺��&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='IsTitlePrefix' onClick=""javascript:PE_TitlePrefix.style.display='';"" value='1' >��&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='IsTitlePrefix' onClick=""javascript:PE_TitlePrefix.style.display='none';"" value='0' checked>��</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TitlePrefix' style='display:none'>" & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5'  align=""right""><strong> ������ǰ׺�ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemTitlePrefix' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>�Ƿ����ݼ�ǰ׺��&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='IsContentPrefix' onClick=""javascript:PE_ContentPrefix.style.display='';"" value='1' >��&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='IsContentPrefix' onClick=""javascript:PE_ContentPrefix.style.display='none';"" value='0' checked>��</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ContentPrefix' style='display:none'>" & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5'  align=""right""><strong> �����ݼ�ǰ׺�ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemContentPrefix' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg""> " & vbCrLf
    Response.Write "           <td colspan=""2"" align=""center"" class=""tdbg"" height=""50"">" & vbCrLf
    Response.Write "            <input name=""Action"" type=""hidden"" id=""Action"" value=""BatchReplace"">" & vbCrLf
    Response.Write "            <input name=""ChannelID"" type=""hidden"" id=""ChannelID"" value=" & ChannelID & ">" & vbCrLf
    Response.Write "            <input name=""Cancel"" type=""button"" id=""Cancel"" value=""������һ��"" onClick=""window.location.href='javascript:history.go(-1)'"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "            <input  type=""submit"" name=""Submit"" value="" ��ʼ�滻 "" onClick=""document.myform.Action.value='DoBatchReplace';"" >" & vbCrLf
    Response.Write "           </td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "       </table>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

'=================================================
'��������DoBatchReplace
'��  �ã������滻����
'=================================================
Sub DoBatchReplace()

    Dim rs, sql, BatchType, BatchArticleID, BatchClassID, ChannelID
    Dim ItemBatchType, ItemReplace, ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult
    Dim ItemBatchTitle, ItemBatchContent, IsTitlePrefix, ItemTitlePrefix, IsContentPrefix, ItemContentPrefix
    Dim FoundErr, ErrMsg

    BatchType = PE_CLng(Trim(Request("BatchType")))
    BatchArticleID = Trim(Request.Form("BatchArticleID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    ChannelID = PE_CLng(Trim(Request.Form("ChannelID")))

    ItemBatchType = PE_CLng(Trim(Request.Form("ItemBatchType")))
    ItemBatchTitle = Trim(Request.Form("ItemBatchTitle"))
    ItemBatchContent = Trim(Request.Form("ItemBatchContent"))
    ItemReplace = Trim(Request.Form("ItemReplace"))
    ItemReplaceStart = Trim(Request.Form("ItemReplaceStart"))
    ItemReplaceEnd = Trim(Request.Form("ItemReplaceEnd"))
    ItemReplaceResult = Trim(Request.Form("ItemReplaceResult"))

    IsTitlePrefix = PE_CLng(Trim(Request.Form("IsTitlePrefix")))
    ItemTitlePrefix = Trim(Request.Form("ItemTitlePrefix"))
    IsContentPrefix = PE_CLng(Trim(Request.Form("IsContentPrefix")))
    ItemContentPrefix = Trim(Request.Form("ItemContentPrefix"))

    If IsTitlePrefix = 0 Then
        ItemTitlePrefix = ""
    Else
        If Len(ItemTitlePrefix) > 100 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ǰ׺�����ܹ���</li>"
        End If
    End If

    If IsContentPrefix = 0 Then
        ItemContentPrefix = ""
    End If

    If BatchType = 1 Then
        If IsValidID(BatchArticleID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����޸ĵ�" & ChannelShortName & "��ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����޸ĵ�" & ChannelShortName & "����Ŀ</li>"
        End If
    End If

    If ItemBatchTitle = "yes" Then
        ItemBatchTitle = True
    End If
    If ItemBatchContent = "yes" Then
        ItemBatchContent = True
    End If

    If ItemBatchTitle = False And ItemBatchContent = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ҫѡ��һ��Ҫ�滻������" & ChannelShortName & "�����" & ChannelShortName & "����</li>"
    End If

    If ItemBatchType = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ��" & ChannelShortName & "�滻�ַ�����</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If ItemBatchType = 1 Then
        If ItemReplace = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ҫ�滻�Ĵ��벻��Ϊ��</li>"
        End If
    ElseIf ItemBatchType = 2 Then
        If ItemReplaceStart = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ҫ�滻�Ŀ�ʼ���벻��Ϊ��</li>"
        End If
        If ItemReplaceEnd = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ҫ�滻��Ľ������벻��Ϊ��</li>"
        End If
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ѡ��" & ChannelShortName & "�滻�ַ����Ͳ���</li>"
    End If

    If ItemReplaceResult = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ҫ�滻��Ĵ��벻��Ϊ��</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If ItemBatchTitle = True Then
        If PE_CLng(Conn.Execute("Select count(*) From PE_Article Where Title='" & ReplaceBadChar(ItemReplaceResult) & "'")(0)) > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>Ҫ�滻�ı��������ݿ����еı����ظ�</li>"
        End If
        If FoundErr = True Then
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If
    End If

    Response.Write "<li>�����滻�������Ժ�..</li>&nbsp;&nbsp;<br>"

    Set rs = Server.CreateObject("ADODB.Recordset")
    If BatchType = 1 Then
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ArticleID in (" & BatchArticleID & ")"
    Else
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ClassID in (" & BatchClassID & ")"
    End If
    rs.Open sql, Conn, 1, 3
    If rs.BOF And rs.EOF Then
        Response.Write "û�п��滻�ı�������ģ�"
    Else
        Do While Not rs.EOF
            If ItemBatchType = 1 Then
                If ItemBatchTitle = True Then
                    If InStr(rs("title"), ItemReplace) <> 0 Then
                        rs("title") = ItemTitlePrefix & Replace(rs("title"), ItemReplace, ItemReplaceResult)
                        Response.Write "<br>&nbsp;&nbsp;" & ChannelShortName & "ID��" & rs("ArticleID") & "&nbsp;&nbsp;" & rs("title") & "..<font color='#009900'>�����滻�ɹ���</font>"
                    End If
                End If
                If ItemBatchContent = True Then
                    If InStr(rs("Content"), ItemReplace) <> 0 Then
                        rs("Content") = ItemContentPrefix & Replace(rs("Content"), ItemReplace, ItemReplaceResult)
                        Response.Write "<br>&nbsp;&nbsp;" & ChannelShortName & "ID��" & rs("ArticleID") & "&nbsp;&nbsp;" & rs("title") & "..<font color='#009900'>�����滻�ɹ���</font>"
                    End If
                End If
            ElseIf ItemBatchType = 2 Then
                If ItemBatchTitle = True Then
                    rs("title") = ItemTitlePrefix & BatchReplaceString(rs("title"), ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult, "����", rs("ArticleID"), rs("title"))
                End If
                If ItemBatchContent = True Then
                    rs("Content") = ItemContentPrefix & BatchReplaceString(rs("Content"), ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult, "����", rs("ArticleID"), rs("title"))
                End If
            End If
            rs("CreateTime") = rs("UpdateTime")

            rs.Update
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Call ClearSiteCache(0)
    Response.Write "<br>&nbsp;&nbsp;<font color='red'>" & ChannelShortName & "�滻�������</font>"
    Response.Write "<br><center>&nbsp;&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>����" & ChannelShortName & "����</a> </center>"
End Sub

'**************************************************
'��������BatchReplaceString
'��  �ã������滻������
'��  ����ItemContent ----����
'��  ����ItemReplaceStart ----���Ҫ�滻�Ŀ�ͷ����
'��  ����ItemReplaceEnd ----���Ҫ�滻�Ľ�������
'��  ����ItemReplaceResult ----Ҫ�滻�Ĵ���
'��  ����ItemName ----��������
'����ֵ��True  ----�Ѵ���
'**************************************************
Function BatchReplaceString(ItemContent, ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult, ItemName, ArticleID, Title)
    If InStr(ItemContent, ItemReplaceStart) = 0 Or InStr(ItemContent, ItemReplaceEnd) = 0 Then
        BatchReplaceString = ItemContent
        Exit Function
    End If
    If GetBody(ItemContent, ItemReplaceStart, ItemReplaceEnd, True, True) = "" Then
        BatchReplaceString = ItemContent
        Exit Function
    End If
    BatchReplaceString = Replace(ItemContent, GetBody(ItemContent, ItemReplaceStart, ItemReplaceEnd, True, True), ItemReplaceResult)
    Response.Write "<br>&nbsp;&nbsp;" & ChannelShortName & "ID��" & ArticleID & "&nbsp;&nbsp;" & Title & "..<font color='#009900'>" & ItemName & "�滻�ɹ���</font>"
End Function


'******************************************************************************************
'����Ϊ���ù̶����Ƽ�������ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub SetProperty()
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    End If
    
    Dim sqlProperty, rsProperty, arrUser
    arrUser = ""
    If InStr(ArticleID, ",") > 0 Then
        sqlProperty = "select * from PE_Article where ArticleID in (" & ArticleID & ")"
    Else
        sqlProperty = "select * from PE_Article where ArticleID=" & ArticleID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        If CheckClassPurview(Action, rsProperty("ClassID")) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�� " & rsProperty("Title") & " û�в���Ȩ�ޣ�</li>"
        Else
            If FoundInArr(arrUser, rsProperty("Inputer"), ",") = False Then
                If arrUser = "" Then
                    arrUser = rsProperty("Inputer")
                Else
                    arrUser = arrUser & "," & rsProperty("Inputer")
                End If
            End If
            Select Case Action
            Case "SetOnTop"
                rsProperty("OnTop") = True
            Case "CancelOnTop"
                rsProperty("OnTop") = False
            Case "SetElite"
                rsProperty("Elite") = True
            Case "CancelElite"
                rsProperty("Elite") = False
            Case "SetPassed"
                If rsProperty("Status") < MyStatus And MyStatus = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp+" & rsProperty("PresentExp") & " where UserName='" & rsProperty("Inputer") & "'")
                End If
                rsProperty("Status") = MyStatus
                If MyStatus < 3 Or CheckLevel = 1 Then
                    rsProperty("Editor") = AdminName
                End If
            Case "CancelPassed", "Reject"
                If rsProperty("Status") = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp-" & rsProperty("PresentExp") & " where UserName='" & rsProperty("Inputer") & "'")
                End If
                If Action = "CancelPassed" Then
                    rsProperty("Status") = 0
                Else
                    rsProperty("Status") = -2
                End If
            End Select
            rsProperty("CreateTime") = rsProperty("UpdateTime")
            rsProperty.Update
        End If
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing

    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)

    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Article.asp?ChannelID=" & ChannelID)

    Call CreateAllJS
End Sub


'******************************************************************************************
'����Ϊ�ƶ�����Ŀ��ר��Ȳ���ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub DoMoveToClass()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim ArticleType, BatchArticleID, BatchClassID
    Dim tChannelID, tClassID, tChannelDir, tUploadDir
    
    ArticleType = PE_CLng(Trim(Request("ArticleType")))
    BatchArticleID = Trim(Request.Form("BatchArticleID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    tChannelID = Trim(Request("tChannelID"))
    tClassID = Trim(Request("tClassID"))
    
    If ArticleType = 1 Then
        If IsValidID(BatchArticleID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ƶ���" & ChannelShortName & "��ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ƶ���" & ChannelShortName & "����Ŀ</li>"
        End If
    End If
    If tChannelID = "" Then
        tChannelID = ChannelID
    Else
        tChannelID = PE_CLng(tChannelID)
        If tChannelID <> ChannelID Then
            If AdminPurview > 1 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>�Բ������Ȩ�޲�����</li>"
            Else
                Dim rsChannel
                Set rsChannel = Conn.Execute("select ChannelDir,UploadDir from PE_Channel where ChannelID=" & tChannelID & "")
                If rsChannel.BOF And rsChannel.EOF Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�Ҳ���Ŀ��Ƶ����</li>"
                Else
                    tChannelDir = rsChannel("ChannelDir")
                    tUploadDir = rsChannel("UploadDir")
                End If
                Set rsChannel = Nothing
            End If
        End If
    End If
    If tClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ŀ����Ŀ������ָ��Ϊ�ⲿ��Ŀ��</li>"
    Else
        tClassID = PE_CLng(tClassID)
        If tClassID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>Ŀ����Ŀ���������" & ChannelShortName & "</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    
    Dim rsBatchMove, sqlBatchMove, ArticlePath
    sqlBatchMove = "select A.ArticleID,A.UploadFiles,A.UpdateTime,A.PaginationType,C.ParentDir,C.ClassDir from PE_Article A left join PE_Class C on A.ClassID=C.ClassID"
    If ArticleType = 1 Then
        sqlBatchMove = sqlBatchMove & " where A.ChannelID=" & ChannelID & " and A.ArticleID in (" & BatchArticleID & ")"
    Else
        sqlBatchMove = sqlBatchMove & " where A.ChannelID=" & ChannelID & " and A.ClassID in (" & BatchClassID & ")"
    End If
    BatchArticleID = ""
    Set rsBatchMove = Conn.Execute(sqlBatchMove)
    Do While Not rsBatchMove.EOF
        ArticlePath = HtmlDir & GetItemPath(StructureType, rsBatchMove("ParentDir"), rsBatchMove("ClassDir"), rsBatchMove("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsBatchMove("UpdateTime"), rsBatchMove("ArticleID"))
        If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
        End If
        If rsBatchMove("PaginationType") > 0 Then
            DelSerialFiles Server.MapPath(ArticlePath) & "_*" & FileExt_Item
        End If
        If tChannelID <> ChannelID Then
            Call MoveUpFiles(rsBatchMove("UploadFiles") & "", tChannelDir & "/" & tUploadDir)    '�ƶ��ϴ��ļ�
        End If
        If BatchArticleID = "" Then
            BatchArticleID = rsBatchMove("ArticleID")
        Else
            BatchArticleID = BatchArticleID & "," & rsBatchMove("ArticleID")
        End If
        rsBatchMove.MoveNext
    Loop
    rsBatchMove.Close
    Set rsBatchMove = Nothing
    If BatchArticleID <> "" Then
        Conn.Execute ("update PE_Article set ChannelID=" & tChannelID & ",ClassID=" & tClassID & ",TemplateID=0,CreateTime=UpdateTime where ArticleID in (" & BatchArticleID & ")")
    End If

    Call WriteSuccessMsg("�ɹ���ѡ����" & ChannelShortName & "�ƶ���Ŀ��Ƶ����Ŀ����Ŀ�У�", "Admin_Article.asp?ChannelID=" & ChannelID & "")
    Call ClearSiteCache(0)
End Sub


Sub MoveUpFiles(strFiles, strTargetDir)
    On Error Resume Next
    Dim strTrueFile, arrFiles, strTrueDir, i
    If IsNull(strFiles) Or strFiles = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    arrFiles = Split(strFiles, "|")
    For i = 0 To UBound(arrFiles)
        strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(arrFiles(i), InStr(arrFiles(i), "/")))
        If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
        strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir & "/" & arrFiles(i))
        If fso.FileExists(strTrueFile) Then
            fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & arrFiles(i))
        End If
    Next
End Sub

'******************************************************************************************
'����Ϊɾ������ա���ԭ�Ȳ���ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub Del()
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, ArticlePath, arrUser
    arrUser = ""
    sqlDel = "select A.ArticleID,A.Title,A.UpdateTime,A.CreateTime,A.Inputer,A.Status,A.Deleted,A.PaginationType,A.PresentExp,A.ReceiveUser,A.ClassID,C.ParentDir,C.ClassDir from PE_Article A left join PE_Class C on A.ClassID=C.ClassID"
    If InStr(ArticleID, ",") > 0 Then
        sqlDel = sqlDel & " where A.ArticleID in (" & ArticleID & ") order by A.ArticleID"
    Else
        sqlDel = sqlDel & " where A.ArticleID=" & ArticleID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        PurviewChecked = False
        ClassID = rsDel("ClassID")
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or (rsDel("Inputer") = UserName And rsDel("Status") = 0) Then
            PurviewChecked = True
        Else
            PurviewChecked = CheckClassPurview(Action, ClassID)
        End If
        
        If PurviewChecked = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ɾ�� <font color='red'>" & rsDel("Title") & "</font> ʧ�ܣ�ԭ��û�в���Ȩ�ޣ�</li>"
        Else
            If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
                If arrUser = "" Then
                    arrUser = rsDel("Inputer")
                Else
                    arrUser = arrUser & "," & rsDel("Inputer")
                End If
            End If
            ArticlePath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("ArticleID"))
            If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
                DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
            End If
            If rsDel("PaginationType") > 0 Then
                DelSerialFiles Server.MapPath(ArticlePath) & "_*" & FileExt_Item
            End If

            If rsDel("Status") = 3 Then
                Conn.Execute ("update PE_User set UserExp=UserExp-" & rsDel("PresentExp") & " where UserName='" & rsDel("Inputer") & "'")
            End If
            rsDel("Deleted") = True
            rsDel("CreateTime") = rsDel("UpdateTime")
            Call Del_User_UnsignedItems(rsDel("ArticleID"), rsDel("ReceiveUser"))
            rsDel.Update
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing

    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)

    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Article.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub

Sub DelFile()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, ArticlePath
    sqlDel = "select A.ArticleID,A.UpdateTime,A.PaginationType,C.ParentDir,C.ClassDir from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where A.ArticleID in (" & ArticleID & ") order by A.ArticleID"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        ArticlePath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("ArticleID"))
        If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
        End If
        If rsDel("PaginationType") > 0 Then
            DelSerialFiles Server.MapPath(ArticlePath) & "_*" & FileExt_Item
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("update PE_Article set CreateTime=UpdateTime where ArticleID in (" & ArticleID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub ConfirmDel()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel
    sqlDel = "select UploadFiles,VoteID from PE_Article where ArticleID in (" & ArticleID & ")"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        Call DelUploadFiles(rsDel("UploadFiles"))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("delete from PE_Article where ArticleID in (" & ArticleID & ")")
    Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
    Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub ClearRecyclebin()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel
    ArticleID = ""
    sqlDel = "select ArticleID,UploadFiles,VoteID from PE_Article where Deleted=" & PE_True & " and ChannelID=" & ChannelID
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        If ArticleID = "" Then
            ArticleID = rsDel(0)
        Else
            ArticleID = ArticleID & "," & rsDel(0)
        End If
        Call DelUploadFiles(rsDel("UploadFiles"))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    If ArticleID <> "" Then
        Conn.Execute ("delete from PE_Article where Deleted=" & PE_True & " and ChannelID=" & ChannelID & "")
        Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
        Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
    End If
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub Restore()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, arrUser
    arrUser = ""
    If InStr(ArticleID, ",") > 0 Then
        sqlDel = "select * from PE_Article where ArticleID in (" & ArticleID & ")"
    Else
        sqlDel = "select * from PE_Article where ArticleID=" & ArticleID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
            If arrUser = "" Then
                arrUser = rsDel("Inputer")
            Else
                arrUser = arrUser & "," & rsDel("Inputer")
            End If
        End If
        If rsDel("Status") = 3 Then
            Conn.Execute ("update PE_User set UserExp=UserExp+" & rsDel("PresentExp") & " where UserName='" & rsDel("Inputer") & "'")
        End If
        rsDel("Deleted") = False
        rsDel.Update
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing

    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)

    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Article.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub

Sub RestoreAll()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, arrUser
    arrUser = ""
    sqlDel = "select * from PE_Article where Deleted=" & PE_True & " and ChannelID=" & ChannelID
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
            If arrUser = "" Then
                arrUser = rsDel("Inputer")
            Else
                arrUser = arrUser & "," & rsDel("Inputer")
            End If
        End If
        If rsDel("Status") = 3 Then
            Conn.Execute ("update PE_User set UserExp=UserExp+" & rsDel("PresentExp") & " where UserName='" & rsDel("Inputer") & "'")
        End If
        rsDel("Deleted") = False
        rsDel.Update
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)
    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Article.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub


Sub DelUploadFiles(strUploadFiles)
    On Error Resume Next
    If Trim(strUploadFiles) = "" Or ObjInstalled_FSO <> True Then Exit Sub
    Dim arrUploadFiles, strFileName, i
    arrUploadFiles = Split(strUploadFiles, "|")
    For i = 0 To UBound(arrUploadFiles)
        If Trim(arrUploadFiles(i)) <> "" Then
            strFileName = InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUploadFiles(i)
            If fso.FileExists(Server.MapPath(strFileName)) Then
                fso.DeleteFile (Server.MapPath(strFileName))
            End If
        End If
    Next
End Sub

Sub Del_User_UnsignedItems(ByVal ArticleID, ByVal strUser)
    If IsNull(strUser) Or strUser = "" Then Exit Sub
    strUser = Replace(strUser, "|", ",")
    strUser = Replace(strUser, ",", "','")
    
    Dim arrID
    arrID = Split(ArticleID, ",")
    Dim sqlUser, rsUser, i, tmpUnsignedItems, tmpArticleID
    Set rsUser = Server.CreateObject("adodb.recordset")
    sqlUser = "select UserID,UserName,UnsignedItems from PE_User where UserName in ('" & strUser & "')"
    rsUser.Open sqlUser, Conn, 1, 3
    Do While Not rsUser.EOF
        For i = 0 To UBound(arrID)
            If FoundInArr(rsUser("UnsignedItems"), CStr(arrID(i)), ",") = True Then
                tmpUnsignedItems = "," & rsUser("UnsignedItems") & ","
                tmpArticleID = "," & arrID(i) & ","
                tmpUnsignedItems = Replace(tmpUnsignedItems, tmpArticleID, ",")
                If tmpUnsignedItems = "," Then
                    rsUser("UnsignedItems") = ""
                Else
                    rsUser("UnsignedItems") = Mid(tmpUnsignedItems, 2, Len(tmpUnsignedItems) - 2)
                End If
                rsUser.Update
            End If
        Next
        rsUser.MoveNext
    Loop
    rsUser.Close
    Set rsUser = Nothing
End Sub

Sub Add_User_UnsignedItems(ByVal ArticleID, ByVal ReceiveUser)
    If IsNull(ReceiveUser) Or ReceiveUser = "" Then Exit Sub
    ReceiveUser = Replace(ReceiveUser, "|", ",")
    ReceiveUser = Replace(ReceiveUser, ",", "','")
    
    Dim sqlUser, rsUser, i
    
    Set rsUser = Server.CreateObject("adodb.recordset")
    sqlUser = "select UserID,UserName,UnsignedItems from PE_User where UserName in ('" & ReceiveUser & "')"
    rsUser.Open sqlUser, Conn, 1, 3
    Do While Not rsUser.EOF
        If rsUser("UnsignedItems") = "" Or IsNull(rsUser("UnsignedItems")) Then
            rsUser("UnsignedItems") = ArticleID
        Else
            If FoundInArr(rsUser("UnsignedItems"), CStr(ArticleID), ",") = False Then
                rsUser("UnsignedItems") = rsUser("UnsignedItems") & "," & ArticleID
            End If
        End If
        rsUser.Update
        rsUser.MoveNext
    Loop
    rsUser.Close
    Set rsUser = Nothing
End Sub

Function UnsignedItemsState(ArticleID)
    Dim rsState, sqlState, strState, arrUser, i
    Dim NotReceiveUser
    sqlState = "select top 1 ReceiveUser,Received from PE_Article where ArticleID=" & ArticleID
    Set rsState = Conn.Execute(sqlState)
    If Not (rsState.BOF And rsState.EOF) Then
        If rsState("Received") = "" Then
            NotReceiveUser = rsState("ReceiveUser")
        Else
            NotReceiveUser = ""
            arrUser = Split(rsState("ReceiveUser"), ",")
            For i = 0 To UBound(arrUser)
                If FoundInArr(rsState("Received"), arrUser(i), "|") = False Then
                    If NotReceiveUser = "" Then
                        NotReceiveUser = arrUser(i)
                    Else
                        NotReceiveUser = NotReceiveUser & "," & arrUser(i)
                    End If
                End If
            Next
        End If
        If NotReceiveUser <> "" Then
            strState = strState & "<a href='' onclick='return false' title='"
            strState = strState & "Ҫ��ǩ���û���" & rsState("ReceiveUser") & vbCrLf
            strState = strState & "�Ѿ�ǩ���û���" & rsState("Received") & vbCrLf
            strState = strState & "��δǩ���û���" & NotReceiveUser
            strState = strState & "'><font color=red>[ǩ����]</font></a>"
        Else
            strState = strState & "<a href='#' title='"
            strState = strState & "ǩ���û���" & rsState("ReceiveUser")
            strState = strState & "'><font color=green>[��ǩ��]</font></a>"
        End If
    End If
    rsState.Close
    Set rsState = Nothing
    UnsignedItemsState = strState
End Function



'*****************************************
'�� �� ����getWordNumber()
'��    ����str �ַ���
'�� �� ֵ����������
'��    �ã��������µ����� ���Լ��㴿���ģ���Ӣ�ģ���Ӣ���ţ���Χ��20������
'��    �ߣ�׳־��������
'�������ڣ�2005-09-07
'*****************************************
Function getWordNumber(ByVal str)
    str = nohtml(PE_HtmlDecode(str))
    regEx.Pattern = "[a-z\-]+|\.+"
    str = regEx.Replace(str, "��")
    str = Replace(str, " ", "")
    getWordNumber = Len(str)
End Function


Sub ExportExcel()
    Dim strSql, SelectType, rsArticleOut, PayStatus, searchDate
    SelectType = Trim(Request("SelectType")) '��������ID���� ����
    strSql = "Select * From PE_Article Where Copymoney>0  And Status=3"
    PayStatus = Trim(Request("PayStatus"))
    If PayStatus = "False" Then
        PayStatus = PE_False
        searchDate = "UpDateTime"
    Else
        PayStatus = PE_True
        searchDate = "PayDate"
    End If
    Select Case SelectType
    Case "ID"
        Dim BeginID, EndID
        BeginID = PE_CLng(Trim(Request("BeginID")))
        EndID = PE_CLng(Trim(Request("EndID")))
        If BeginID <> 0 And EndID <> 0 Then
            strSql = strSql & "And Ispayed=" & PayStatus & " and (ArticleID Between " & BeginID & " and " & EndID & ")"
        End If
    Case "Date"
        Dim BeginDate, EndDate
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        If BeginDate = "" Then
            BeginDate = "1900-1-1"
        Else
            BeginDate = ReplaceBadChar(BeginDate)
        End If
        If EndDate = "" Then
            EndDate = FormatDateTime(Date, 2)
        Else
            EndDate = ReplaceBadChar(EndDate)
        End If
        If IsDate(BeginDate) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��������ȷ����ʼ���ڣ�</li>"
        End If
        If IsDate(EndDate) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��������ȷ�Ľ������ڣ�</li>"
        End If
        If FoundErr = True Then
            Exit Sub
        End If
        If SystemDatabaseType = "SQL" Then
            strSql = strSql & "And Ispayed=" & PayStatus & " and UpdateTime Between '" & BeginDate & "' and '" & EndDate & "'"
        Else
            strSql = strSql & "And Ispayed=" & PayStatus & " and UpdateTime Between #" & BeginDate & "# and #" & EndDate & "#"
        End If
    Case Else
        If InStr(ArticleID, ",") >= 0 And ArticleID <> "" Then
            strSql = strSql & "And Ispayed=" & PayStatus & " And ArticleID in (" & ArticleID & ")"
        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
            Exit Sub
        End If
    End Select
    Set rsArticleOut = Conn.Execute(strSql)
    If rsArticleOut.BOF And rsArticleOut.EOF Then
        Response.Write "<script language='javascript'>alert('û�в�ѯ���������')</script>"
    Else
        Call outHead2
        Response.Write "<table border=""0"" cellspacing=""1"" style=""border-collapse: collapse;table-layout:fixed"" id=""AutoNumber1"" height=""32"">" & vbCrLf
        Response.Write "<tr>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>����ID</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>���±���</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>¼����</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>����</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>���������</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>��������</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>���(��λ��Ԫ)</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>��֧��</b></span></td>" & vbCrLf
        If PayStatus = PE_True Then
            Response.Write "<td align=""center""><span lang=""zh-cn""><b>֧��ʱ��</b></span></td>" & vbCrLf
        Else
            Response.Write "<td align=""center""><span lang=""zh-cn""><b>¼��ʱ��</b></span></td>" & vbCrLf
        End If
        Response.Write "</tr>" & vbCrLf
        Do While Not rsArticleOut.EOF
            Response.Write "<tr>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("ArticleID") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Title") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Inputer") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Author") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Beneficiary") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & getWordNumber(rsArticleOut("Content")) & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Copymoney") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">"
            If rsArticleOut("IsPayed") Then
                Response.Write "��"
            Else
                Response.Write "��"
            End If
            Response.Write "</span></td>" & vbCrLf
            If PayStatus = PE_True Then
                Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("payDate") & "</span></td>" & vbCrLf
            Else
                Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("UpdateTime") & "</span></td>" & vbCrLf
            End If
            Response.Write "</tr>" & vbCrLf
            rsArticleOut.MoveNext
        Loop
        rsArticleOut.Close
        Set rsArticleOut = Nothing
        Response.Write "</table>" & vbCrLf
    End If
End Sub

Sub ConfirmPay()
   Dim strSql, rsArticle, i, arrArticleID
   strSql = "Update PE_Article Set IsPayed=" & PE_True & ",PayDate=" & PE_Now & " Where Copymoney>0 And Ispayed=" & PE_False & " And Status=3"
   If InStr(ArticleID, ",") >= 0 And ArticleID <> "" Then
      strSql = strSql & " And ArticleID in (" & ArticleID & ")"
   Else
      FoundErr = True
      ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
      Exit Sub
   End If
   Conn.Execute (strSql)
   Conn.Close
   Set Conn = Nothing
   Response.Redirect "Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=PayMoney&Status=9"
End Sub

Sub outHead2()
    Response.Write "<html><head>" & vbCrLf
    Response.ContentType = "application/vnd.ms-excel" & vbCrLf
    Response.AddHeader "Content-Disposition", "attachment"
    Response.Write "<meta http-equiv=""Content-Language"" content=""zh-cn"">" & vbCrLf
    Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
    Response.Write "<title>����б�</title>" & vbCrLf
    Response.Write "<body>" & vbCrLf
End Sub


%>
