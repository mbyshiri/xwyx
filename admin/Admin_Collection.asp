<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Collection.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
<!--#include file="../Include/PowerEasy.CreateThumb.asp"-->

<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Collection"   '����Ȩ��

Private rs, sql, rsItem, i 'ͨ�ñ���
'��Ŀ���ñ���
Private ItemID, ItemName, ClassID, SpecialID, ItemNum, ListNum, ItemEnd, ListEnd
Private Status, OnTop, Elite, Hot, Hits, Stars, InfoPoint, MaxCharPerPage, ShowCommentLink
Private FilterProperty, Script_Iframe, Script_Object, Script_Script, Script_Class, Script_Div, Script_Span, Script_Img, Script_Font, Script_A, Script_Html, Script_Table, Script_Tr, Script_Td
Private SaveFiles, SaveFlashUrlToFile, CollecOrder, CreateImmediate
Private ListStr, LsString, LoString, ListPaingType, LPsString, LPoString, ListPaingStr1, ListPaingStr2, ListPaingID1, ListPaingID2, ListPaingStr3, HsString, HoString, HttpUrlType, HttpUrlStr
Private TsString, ToString, CsString, CoString, AuthorType, AsString, AoString, AuthorStr, CopyFromType, FsString, FoString
Private CopyFromStr, KeyType, KsString, KoString, KeyStr, KeyScatterNum, NewsPaingType, NPsString, NPoString, NewsPaingStr1, NewsPaingStr2
Private PsString, PoString, PhsString, PhoString
Private IsString, IoString, IntroType, IntroStr, IntroNum, Intro
Private DateType, DsString, DoString
Private IncludePicYn, DefaultPicYn, PaginationType
'�Զ����ֶβɼ�����
Private IsField, Field, iField
Private arrField, arrField2, FieldID, FieldName, FieldType, FisSting, FioSting, FieldStr
'��¼��֤����
Private LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse
'�ɼ�ѡ�����
Private CollecTest, Content_view, IsTitle, IsLink
'�ɼ���صı���
Private CollecNewsi, CollecNewsj, CollectionModify, ItemIDStr
Private ItemIDArray, ContentTemp, NewsPaingNextCode, NewsPaingNext
Private Arr_j, Arr_i, NewsUrl, NewsCode
Private LoginData, LoginResult, CollecNewsA, OrderTemp, StartTime
'ͼƬ���ͼ�����·��
Private FilesOverStr, FilesPath, FilesArray, ImagesNum
'���±������
Private ArticleID, Title, Content, Author, CopyFrom, Key, UpDateType, UpdateTime, IncludePic, UploadFiles, DefaultPicUrl
'��ʷ��¼
Private His_Title, His_NewsCollecDate, His_Result, His_Repeat, His_i
'�ɼ��б������
Private WebUrl, ListUrl, ListCode, ListUrlArray, NewsArrayCode, NewsArray, ListArray, ListPaingNext
Private tempStr, ItemIDtemp, TimeNum, rnd_temp, ArticleList, CollectionNum, CollectionType
Private AddWatermark, AddThumb, ItemSucceedNum, ItemSucceedNum2, ImagesNumAll, PaingNum, dirMonth, dtNow
Private Arr_Item, Arr_Histrolys, CollecType, Arr_Filters, Filteri, FilterStr, SwfTime, CollectionCreateHTML '�ɼ�����
'�ɼ����ķ�ҳ����
Private PageListCode, PageArrayCode, PageArray
'��ʱ����
Private Timing_AreaCollection, TimingCreate
'�շ���������
Private InfoPurview, arrGroupID, PitchTime, ReadTimes, DividePercent
'�б�����ͼ
Private ThumbnailType, ThsString, ThoString
Private ThumbnailArrayCode, ThumbnailArray, ThumbnailUrl
'ת��·��
Private ConversionTrails


'�ɼ���ʱˢ�»��������http://
If InStr(ComeUrl, "?") > 0 Then
    ComeUrl = Left(ComeUrl, InStr(ComeUrl, "?"))
End If

    
'��õ�ǰʱ�䵱ǰ����
dtNow = Now()
dirMonth = Year(dtNow) & Right("0" & Month(dtNow), 2)


XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))



Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>�ɼ�ϵͳ</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf

'��Щ������ʱ�ŵ��ɼ��Ժ���ƿ�
If Action = "CreateItemHtml" Then
Else
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""border"">" & vbCrLf
    Call ShowPageTitle("�� �� ϵ ͳ �� Ŀ �� ��", 10051)
    Response.Write "</table>" & vbCrLf
End If

Select Case Action
Case "Start"                    '��ʼ�ɼ�
    Call Start
Case "main"                     '�ɼ�����
    Call main
Case "CheckItem"
    Call CheckItem              '���������Ŀ
Case "StopCollection"           'ֹͣ�ɼ�����
    ItemEnd = True
    CollecNewsi = PE_CLng(Trim(Request("CollecNewsi")))         'CollecNewsi    ��ʾ�ɼ��ɹ���
    CollecNewsj = PE_CLng(Trim(Request("CollecNewsj")))         'CollecNewsj    ��ʾ�ɼ�ʧ����
    ArticleList = Replace(CStr(Trim(Request("ArticleList"))),"|","/")            'ArticleList ���ڻ��治ͬ���б�
    ItemSucceedNum2 = PE_CLng(Trim(Request("ItemSucceedNum2"))) 'ItemSucceedNum2 �ɹ��ɼ���Ŀ��
    ImagesNumAll = PE_CLng(Trim(Request("ImagesNumAll")))       'ImagesNumAll    ��Ŀ����
    CollecType = PE_CLng(Trim(Request("CollecType")))           'CollecType    �ɼ�ģʽ 0 �ȶ� 1 ����
    CollectionCreateHTML = Trim(Request("CollectionCreateHTML")) 'CollectionCreateHTML    ����html����
    CreateImmediate = Trim(Request("CreateImmediate"))          'CreateImmediate �ɼ���Ŀ�Ƿ�����
    UseCreateHTML = PE_CLng(Trim(Request("UseCreateHTML")))     'UseCreateHTML   Ƶ���Ƿ�����

    If CollectionCreateHTML = "" Then
        If CreateImmediate = "True" And UseCreateHTML <> 0 And ItemSucceedNum2 <> 0 Then
            CollectionCreateHTML = PE_CLng(Trim(Request("ChannelID"))) & "$" & PE_CLng(Trim(Request("ClassID"))) & "$" & ReplaceBadChar(Trim(Request("SpecialID"))) & "$" & ItemSucceedNum2
        End If
    Else '����Ƕ���Ŀ��վֹͣ

        If CreateImmediate = "True" And UseCreateHTML <> 0 And ItemSucceedNum2 <> 0 Then
            CollectionCreateHTML = CollectionCreateHTML & "|" & PE_CLng(Trim(Request("ChannelID"))) & "$" & PE_CLng(Trim(Request("ClassID"))) & "$" & ReplaceBadChar(Trim(Request("SpecialID"))) & "$" & ItemSucceedNum2
        End If
    End If

    ErrMsg = "<br>�Ѿ�ֹͣ��ǰ��Ŀ,Ŀǰ����ɣ�"
    ErrMsg = ErrMsg & "<li>�ɹ��ɼ��� <font color=red>" & CollecNewsi & "</font>  ƪ,ʧ�ܣ�<font color=blue> " & CollecNewsj & "</font>  ƪ,ͼƬ��<font color=green>" & ImagesNumAll & "</font> ����</li>"
    Call PE_Cache.DelAllCache
    Call WriteSuccessMsg2(ErrMsg)
Case "CreateItemHtml"
    Call CreateItemHtml         '�ɼ����Զ�����Html
Case Else
    Call main
End Select
Response.Write "</body></html>"
Call CloseConn


'=================================================
'��������Main
'��  �ã����²ɼ�
'=================================================
Sub main()

    Dim sql, rs, SqlH, RsH, Flag, Action
    Dim iChannelID, ClassID, SpecialID, ItemID, ItemName, ListUrl, WebName, NewsCollecDate
    Dim SkinID, LayoutID, SkinCount, LayoutCount, MaxPerPage
        
    If Request("page") <> "" Then
        CurrentPage = CInt(Request("page"))
    Else
        CurrentPage = 1
    End If
    iChannelID = PE_CLng(Trim(Request("iChannelID")))
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 10
    
    strFileName = "Admin_Collection.asp?Action=Main&iChannelID=" & iChannelID

    Response.Write "<a name='submit'></a>" & vbCrLf
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""70"" height=""30""><strong>��������</strong></td>" & vbCrLf
    Response.Write "    <td height=""30""><a href=Admin_Collection.asp?Action=Main>������ҳ</a> | <a href=""Admin_CollectionManage.asp?Action=Step1"">�������Ŀ</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
    
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "    function CheckAll(thisform){" & vbCrLf
    Response.Write "        for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "            var e = thisform.elements[i];" & vbCrLf
    Response.Write "            if (e.Name != ""chkAll""&&e.disabled!=true&&e.zzz!=1)" & vbCrLf
    Response.Write "                e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function mysub(){" & vbCrLf
    Response.Write "        window.location='#submit';" & vbCrLf
    Response.Write "        esave.style.visibility=""visible"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    
    If IsObjInstalled("MSXML2.XMLHTTP") = False Then
        Call WriteErrMsg("<li>����ϵͳû�а�װXMLHTTP ���,�뵽΢����վ����MSXML 4.0", ComeUrl)
        Exit Sub
    End If

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "<tr class='title'><td colspan='2'> | "
    sql = "SELECT DISTINCT I.ChannelID, C.ChannelName,C.ModuleType FROM PE_Item I LEFT OUTER JOIN PE_Channel C ON I.ChannelID = C.ChannelID"
    sql = sql & " WHERE C.ModuleType=1"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
    Else
        Do While Not rs.EOF
            Response.Write "<a href='Admin_Collection.asp?Action=Main&iChannelID=" & rs("ChannelID") & "'><FONT style='font-size:12px'"
            If rs("ChannelID") = iChannelID Then Response.Write "color='red'"
            Response.Write ">" & rs("ChannelName") & "</FONT></a> | "
            rs.MoveNext
        Loop
        Response.Write "<a href='Admin_Collection.asp?Action=Main&iChannelID=0'><FONT style='font-size:12px'"
        If iChannelID = 0 Then Response.Write "color='red'"
        Response.Write "> ����Ƶ�� </FONT></a> | "
    End If
    Response.Write "</td></tr>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
    Response.Write "<br>"
    Response.Write GetManagePath(iChannelID)
    Response.Write "<br>"
    Response.Write "<table class=""border"" border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">" & vbCrLf
    Response.Write "<form name=""myform"" method=""POST"" action=""Admin_Collection.asp"">" & vbCrLf
    Response.Write "  <tr class=""title"" style=""padding: 0px 2px;"">" & vbCrLf
    Response.Write "    <td width=""40"" height=""22"" align=""center""><strong>ѡ��</strong></td>        " & vbCrLf
    Response.Write "    <td width=""100"" align=""center""><strong>��Ŀ����</strong></td>" & vbCrLf
    Response.Write "    <td width=""100"" align=""center""><strong>�ɼ���ַ</strong></td>" & vbCrLf
    Response.Write "    <td width=""100"" height=""22"" align=""center""><strong>����Ƶ��</strong></td> " & vbCrLf
    Response.Write "    <td width=""100"" height=""22"" align=""center""><strong>������Ŀ</strong></td> " & vbCrLf
    Response.Write "    <td width=""40"" align=""center""><strong>������</strong></td>        " & vbCrLf
    Response.Write "    <td width=""120"" height=""22"" align=""center""><strong>�ϴβɼ�ʱ��</strong></td>" & vbCrLf
    Response.Write "    <td width=""60"" height=""22"" align=""center""><strong>�ɹ���¼</strong></td>" & vbCrLf
    Response.Write "    <td width=""60"" height=""22"" align=""center""><strong>ʧ�ܼ�¼</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    sql = "SELECT I.*,C.ChannelName,CL.ClassName,C.Disabled,C.ModuleType "
    sql = sql & " FROM (PE_Item I left JOIN PE_Channel C ON I.ChannelID =C.ChannelID)"
    sql = sql & " Left JOIN PE_Class CL ON I.ClassID = CL.ClassID"
    sql = sql & " where C.ModuleType=1 and I.Flag=" & PE_True
    If iChannelID <> 0 Then sql = sql & " And I.ChannelID=" & iChannelID
    sql = sql & " ORDER BY I.Flag "
        If SystemDatabaseType = "SQL" Then
        sql = sql & "desc"
    Else
        sql = sql & "asc"
    End If
    sql = sql & ", I.ItemID DESC, I.NewsCollecDate DESC"

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg' height='50'><td colspan='9' align='center'>ϵͳ�����޲ɼ���Ŀ��</td></tr></table>"
    Else
        totalPut = rs.RecordCount
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
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rs.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Dim VisitorNum
        VisitorNum = 0
        Do While Not rs.EOF
            ChannelID = rs("ChannelID")
            ClassID = PE_CLng(rs("ClassID"))
            ItemID = rs("ItemID")
            ItemName = rs("ItemName")
            ListUrl = rs("ListStr")
            WebName = rs("WebName")
            NewsCollecDate = rs("NewsCollecDate")
            Flag = rs("Flag")
            
            Response.Write "<tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"" height='25'>" & vbCrLf

            Response.Write "    <input type=""checkbox"" value=" & ItemID & " name=""ItemID"""
            If rs("Disabled") = True Or Flag <> True Or IsNull(rs("ChannelName")) = True Then
                Response.Write " disabled"
            End If
            Response.Write "> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">" & ItemName & "</td> " & vbCrLf
            Response.Write "  <td width=""100"" align=""center""><a href=" & ListUrl & " target=""_bank"">" & WebName & "</a></td>  " & vbCrLf
            Response.Write "  <td width=""100"" height=""22"" align=""center"">"
            If IsNull(rs("ChannelName")) = True Then
                Response.Write "��û��ָ��Ƶ��"
            Else
                If rs("Disabled") = True Then
                    Response.Write rs("ChannelName") & "<font color=red>&nbsp;�ѽ���</font>"
                Else
                    Response.Write rs("ChannelName")
                End If
            End If
            Response.Write "</td> " & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">"
            If IsNull(rs("ClassName")) = True Then
                Response.Write "��û��ָ����Ŀ"
            Else
                Response.Write rs("ClassName")
            End If
            Response.Write "</td>" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"">" & vbCrLf
            If Flag = True Then
                Response.Write "<b>��</b>"
            Else
                Response.Write "<FONT color='red'><b>��</b></FONT>"
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""120"" align=""center"">" & vbCrLf
            If DateDiff("d", NewsCollecDate, Now()) = 0 Then
                Response.Write "<font color=red>" & NewsCollecDate & "</font>"
            Else
                Response.Write NewsCollecDate
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""60"" align=""center"">" & vbCrLf
            Response.Write "   <a href='Admin_CollectionHistory.asp?Action=main&SelectCollateItemID=" & ItemID & "&HistrolyResult=true'>&nbsp;"
            Call HistrolyNum(ItemID, PE_True)
            Response.Write "</a>" & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""60"" align=""center"">" & vbCrLf
            Response.Write "   <a href='Admin_CollectionHistory.asp?Action=main&SelectCollateItemID=" & ItemID & "&HistrolyResult=false'>&nbsp;"
            Call HistrolyNum(ItemID, PE_False)
            Response.Write "</a>" & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "</tr> " & vbCrLf

            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "<tr class='tdbg'>" & vbCrLf
        Response.Write "  <td colspan='7' align=""right"">�ϼƣ�</td>" & vbCrLf
        Response.Write "  <td align='center' width='60'>" & vbCrLf
        SqlH = "select count(HistrolyNewsID) from PE_HistrolyNews where  Result=" & PE_True
        Set RsH = Conn.Execute(SqlH)
        If RsH.BOF And RsH.EOF Then
            Response.Write "&nbsp;<font color='green'>0</font>"
        Else
            Response.Write "&nbsp;<font color='blue'>" & RsH(0) & "</font>"
        End If
        RsH.Close
        Set RsH = Nothing
        Response.Write "  </td>" & vbCrLf
        Response.Write "  <td align='center' width='60'>" & vbCrLf
        SqlH = "select count(HistrolyNewsID) from PE_HistrolyNews where  Result=" & PE_False
        Set RsH = Conn.Execute(SqlH)
        If RsH.BOF And RsH.EOF Then
            Response.Write "&nbsp;<font color='green'>0</font>"
        Else
            Response.Write "&nbsp;<font color='red'>" & RsH(0) & "</font>"
        End If
        RsH.Close
        Set RsH = Nothing
        Response.Write "  </td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf
        Response.Write "<tr class='tdbg'>" & vbCrLf
        Response.Write "  <td colspan='2'>" & vbCrLf
        Response.Write "     <input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" > &nbsp;ȫѡ &nbsp;&nbsp;"
        Response.Write "  </td>" & vbCrLf
        Response.Write "  <td colspan='7'>" & vbCrLf
        Response.Write "    <input name=""ItemNum"" type=""hidden"" id=""ItemNum"" value=""1"">"
        Response.Write "    <input name=""ListNum"" type=""hidden"" id=""ListNum"" value=""1"">"
        Response.Write "    <input name=""Arr_i"" type=""hidden"" id=""Arr_i"" value=""0"">"
        Response.Write "    <input name=""CollecNewsA"" type=""hidden"" id=""CollecNewsA"" value=""0"">"
        Response.Write "    <input name=""CollecNewsi"" type=""hidden"" id=""CollecNewsi"" value=""0"">"
        Response.Write "    <input name=""ItemSucceedNum"" type=""hidden"" id=""ItemSucceedNum"" value=""0"">"
        Response.Write "    <input name=""ItemSucceedNum2"" type=""hidden"" id=""ItemSucceedNum2"" value=""0"">"
        Response.Write "    <input name=""CollecNewsj"" type=""hidden"" id=""CollecNewsj"" value=""0"">"
        Response.Write "    <input name=""ImagesNumAll"" type=""hidden"" id=""ImagesNumAll"" value=""0"">"
        Response.Write "    <input name=""ItemIDtemp"" type=""hidden"" id=""ItemIDtemp"" value=""0"">"
        Response.Write "    <input name=""Action"" type=""hidden"" id=""Action"" value=""Start"">"
        Response.Write "    <input name=""CollecType"" type=""hidden"" id=""ItemNum"" value=""1"">"
        Response.Write "    <INPUT TYPE='checkbox' NAME='CollecTest' value='yes' zzz='1' onclick=""javascript:document.myform.Content_View.checked=true""> ��¼�����ݿ⣬ֻ���Բɼ������Ƿ�����<br>" & vbCrLf
        Response.Write "    <INPUT TYPE='checkbox' NAME='Content_View' value='yes' zzz='1'> �ɼ�������Ԥ����������<br>" & vbCrLf
        Response.Write "    <INPUT TYPE='checkbox' NAME='IsTitle' value='yes' zzz='1'> ���ɼ�������Ŀ�����е���ͬ��������<br>" & vbCrLf
        Response.Write "    <INPUT TYPE='checkbox' NAME='IsLink' value='yes' zzz='1'> �ڲ����Ӳɼ�����ѡ��ֻ������Ӳɼ���<br>" & vbCrLf
        Response.Write "  </td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf
        Response.Write "<tr class='tdbg'>" & vbCrLf
        Response.Write "  <td colspan='9' height='32' align='center'>" & vbCrLf
        Response.Write "    <input type=""submit"" value=""�� �� �� ��"" name=""submit"" onclick=""javascript:mysub();document.myform.Action.value='Start';document.myform.CollecType.value=1"" >&nbsp;&nbsp;&nbsp;"
        Response.Write "    <input type=""submit"" value=""�� �� �� ��"" name=""submit"" onclick=""javascript:mysub();document.myform.Action.value='Start';document.myform.CollecType.value=0"" >&nbsp;&nbsp;&nbsp;"
        Response.Write "    <input type=""submit"" value=""�� �� �� ��"" name=""submit"" onclick=""javascript:if (confirm('���Ӳɼ�������ֻ�ɼ��Է���վ�����ӣ����ɼ����ģ����ｨ�������úòɼ���Ŀ�ı���ͼ�飬�ڰ�Ť���Ϸ����������ڲ����ӻ����ⲿ���ӣ��ڲ����Ӿ�����������ֻ����Է���URL��������ģ�������ҳ������չ���ⲿ���Ӿ����б�����ת�����ӣ���ȷ��ʹ�����Ӳɼ�ô��')){mysub();document.myform.Action.value='Start';document.myform.CollecType.value=2;}else{return false;};"" >&nbsp;&nbsp;&nbsp;"
        Response.Write "    <input type=""submit"" value=""�� �� �� �� "" name=""submit"""
        '�õ��ϵ��¼
        Dim rsBreakpoint
        sql = "select top 1 Timing_Breakpoint from PE_config"
        Set rsBreakpoint = Server.CreateObject("adodb.recordset")
        rsBreakpoint.Open sql, Conn, 1, 3
        If rsBreakpoint("Timing_Breakpoint") = "" Then
            Response.Write " disabled"
        End If
        Response.Write "    onclick=""javascript:if (confirm('�ϴβɼ���Ϊ��ֹͣ�˲ɼ���Ŀ��XMLHTTP������������ϵ�����ֹ���������Ƿ�����ϴεĲɼ���Ŀ��')){mysub();document.myform.Action.value='Start';document.myform.CollecType.value=3;}else{return false;};"" >&nbsp;&nbsp;&nbsp;"
        rsBreakpoint.Close
        Set rsBreakpoint = Nothing
        Response.Write "    <input type=""submit"" value=""���ɼ���Ŀ"" name=""CheckItem"" onclick=""javascript:if (confirm('����Ĳɼ���Ŀ�Ƚ϶࣬���ҳ�ʱ��δʹ�òɼ�ʱ������ܲ���ȷ����Щ�ɼ���Ŀ��������ʹ�ã��ڴ�����������ʹ�ñ���������⡣�˹��ܷǳ���ʱ���뾡�����á�ȷ��Ҫ���м����')){mysub();document.myform.Action.value='CheckItem'}else{return false;};"" ></td>"
        Response.Write "  </td></tr>" & vbCrLf
        Response.Write "</form>" & vbCrLf
        Response.Write "</table>" & vbCrLf

        If totalPut > 0 Then
            Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����Ŀ��¼", True)
        End If
        Response.Write "<br>" & vbCrLf
        Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "   <td colspan='2' height=""20"" align=""center""><font color=#ff6600><strong>��������ʹ�ñ�ϵͳ�ṩ�Ĳɼ�������������µ�һ�з��ɻ򾭼����ζ���ʹ���߳е�����ϵͳ�����̲��е��κ����Σ�</strong></font></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
    rs.Close
    Set rs = Nothing
    Response.Write " <div id=""esave"" style=""position:absolute; top:50px; left:200px; z-index:1;visibility:hidden""> " & vbCrLf
    Response.Write "    <TABLE WIDTH=400 BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
    Response.Write "      <TR><td width=""20%""></td>" & vbCrLf
    Response.Write "    <TD width=""60%""> " & vbCrLf
    Response.Write "    <TABLE WIDTH=100% height=100 BORDER=0 CELLSPACING=1 CELLPADDING=0>" & vbCrLf
    Response.Write "    <TR> " & vbCrLf
    Response.Write "      <td bgcolor=""#0033FF"" align=center><b><marquee align=""middle"" behavior=""alternate"" scrollamount=""5""><font color=#FFFFFF>���ڼ��زɼ���Ŀ,���Ժ�...</font></marquee></b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "    </td><td width='20%'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "  </div>" & vbCrLf
    Response.Write " <table WIDTH=400 height=130 BORDER=0 CELLSPACING=0 CELLPADDING=0><tr><td></td></tr></table>" & vbCrLf
    Call CloseConn
End Sub

'=================================================
'��������Start
'��  �ã����������ɼ�����
'=================================================
Sub Start()
    FoundErr = False                    '�Ƿ��д���
    ItemEnd = False                     '�Ƿ�ɼ���Ŀ���
    ListEnd = False                     '�Ƿ�ɼ��б����
    ErrMsg = ""                         '����˵��
    TimeNum = 3     '�ȴ�ʱ��
    ItemNum = PE_CLng(Trim(Request("ItemNum")))                     'ItemNum    ��Ŀ��
    ListNum = PE_CLng(Trim(Request("ListNum")))                     'ListNum    �б���
    Arr_i = PE_CLng(Trim(Request("Arr_i")))                         'Arr_i      ��ǰ�б�ĵڼ�������
    CollecNewsi = PE_CLng(Trim(Request("CollecNewsi")))             'CollecNewsi    ��ʾ�ɼ��ɹ���
    CollecNewsj = PE_CLng(Trim(Request("CollecNewsj")))             'CollecNewsj    ��ʾ�ɼ�ʧ����
    ListPaingNext = Replace(Trim(Request("ListPaingNext")),"|","/") 'ListPaingNext  ��ʾ�ɼ��б���һҳ
    ItemIDStr = Replace(ReplaceBadChar(Trim(Request("ItemID"))), " ", "") 'ItemIDStr ��Ŀ����
    CollecNewsA = PE_CLng(Trim(Request("CollecNewsA")))             'CollecNewsA �ɼ�������
    ItemIDtemp = PE_CLng(Trim(Request("ItemIDtemp")))               'ItemIDtemp  ��Ŀ�Ƿ��״μ���
    rnd_temp = CStr(Trim(Request("rnd_temp")))                      'rnd_temp    �����������ͬ�Ļ���
    ArticleList = Replace(CStr(Trim(Request("ArticleList"))),"|","/")                      'ArticleList ���ڻ��治ͬ���б�
    ItemSucceedNum = PE_CLng(Trim(Request("ItemSucceedNum")))       'ItemSucceedNum ��Ŀ�ɼ��ɹ���Ϊ��¼��ͬ��Ŀ�ɼ��ɹ������ڶ���Ŀָ���Ĳɼ�����
    ItemSucceedNum2 = PE_CLng(Trim(Request("ItemSucceedNum2")))     'ItemSucceedNum2 �ɹ��ɼ���Ŀ��
    ImagesNumAll = PE_CLng(Trim(Request("ImagesNumAll")))           'ImagesNumAll    ��Ŀ����
    CollecType = PE_CLng(Trim(Request("CollecType")))               'CollecType    �ɼ�ģʽ 0 �ȶ� 1 ���� 2 ���� 3 �ϵ�����
    CollectionCreateHTML = Trim(Request("CollectionCreateHTML"))    'CollectionCreateHTML    ����html����
    TimingCreate = Trim(Request("TimingCreate"))                    'TimingCreate  ��ʱ����html
    Timing_AreaCollection = Trim(Request("Timing_AreaCollection"))  'Timing_AreaCollection  ��ʱ����ɼ�

    If CollecType = 3 Then
        '�ϵ�����
        sql = "select top 1 Timing_Breakpoint from PE_config"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 3
        Response.Write rs("Timing_Breakpoint")
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If Trim(Request("CollecTest")) = "yes" Then
        CollecTest = True
    Else
        CollecTest = False
    End If
    If Trim(Request("Content_view")) = "yes" Then
        Content_view = True
    Else
        Content_view = False
    End If

    If Trim(Request("IsTitle")) = "yes" Then
        IsTitle = True
    Else
        IsTitle = False
    End If

    If Trim(Request("IsLink")) = "yes" Then
        IsLink = True
    Else
        IsLink = False
    End If
	If IsValidID(ItemIDStr) = False Then
		ItemIDStr = ""
	End If

    If ItemIDStr = "" Then
        FoundErr = True
        ErrMsg = "<li>��������,��ѡ����Ŀ��</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    ElseIf ItemIDStr = "0" Then 'Ϊ��ʱ������ת
        Call Refresh("Admin_Timing.asp?Action=DoTiming&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,0)	
        'Response.Write " <meta http-equiv=""refresh"" content=0;url=""Admin_Timing.asp?Action=DoTiming&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
        Exit Sub
    End If
     
    '�Ƿ�ȫ���ɼ����
    'ItemNum ��ǰ��Ŀ�� ItemIDStr ��Ŀ���� �ָ����� ItemIDArray �õ���Ŀ�� ItemIDArray �õ�ÿһ����Ŀ��
    ItemIDArray = Split(ItemIDStr, ",")
    If (ItemNum - 1) > UBound(ItemIDArray) Then
        ItemEnd = True
        ErrMsg = "<br>ȫ����Ŀ�ɼ�������ɣ�"
        ErrMsg = ErrMsg & "<li>�ɹ��ɼ��� <font color=red>" & CollecNewsi & "</font>  ƪ,ʧ�ܣ�<font color=blue> " & CollecNewsj & "</font>  ƪ,ͼƬ��<font color=green>" & ImagesNumAll & "</font> ����</li>"

        '��նϵ��¼
        sql = "select Timing_Breakpoint from PE_config"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 3
        rs("Timing_Breakpoint") = ""
        rs.Update
        rs.Close
        Set rs = Nothing
        
        '�������
        Call PE_Cache.DelAllCache
        Call WriteSuccessMsg2(ErrMsg)
        Exit Sub
    End If
    
    '���س�ʼ��Ŀ�뻺��
    If ItemIDtemp = 0 Then
        Call SetCache
        ItemIDtemp = 1
    Else
        If PE_Cache.CacheIsEmpty("Collection" & rnd_temp) Then
            Call SetCache
            ArticleList = ""
        End If
    End If
    
    '���ػ���
    Arr_Item = PE_Cache.GetValue("Collection" & rnd_temp)
    Arr_Filters = PE_Cache.GetValue("Arr_Filters" & rnd_temp)
    Arr_Histrolys = PE_Cache.GetValue("Arr_Histrolys" & rnd_temp)
    Call loadItem
    
    If CollectionNum <> "" Then   '�Ƿ���ָ���ĳɹ��ɼ���
        If CollectionType = 0 Then  '�Ƿ���ָ���ĳɹ��ɼ���
            If ItemSucceedNum = PE_CLng(CollectionNum) Then
                ErrMsg = "<li>�Ѿ��ɹ��ɼ���" & ItemName & "��Ŀ<font color=red>" & CollectionNum & "</font>ƪָ���ɼ�����</li>"
                ErrMsg = ErrMsg & "<br><font color=red>" & ItemName & "</font> ��Ŀ�ɼ�������ɣ�</li>"
                Call WriteSuccessMsg2(ErrMsg)
                Exit Sub
            End If
        End If
        If CollectionType = 1 Then  '�Ƿ���ÿҳ��Ҫ�Ĳɼ���
            If ListNum > PE_CLng(CollectionNum) Then
                ErrMsg = "<li>�Ѿ��ɹ��ɼ���" & ItemName & "��Ŀ<font color=red>" & CollectionNum & "</font>ƪָ��������</li>"
                ErrMsg = ErrMsg & "<br><font color=red>" & ItemName & "</font> ��Ŀ�ɼ�������ɣ�</li>"
                Call WriteSuccessMsg2(ErrMsg)
                Exit Sub
            End If
        End If
    End If
    
    '������Ŀ��¼ʱ��
    If ListNum = 1 And CollecTest = False Then
        sql = "select top 1 * from PE_Item where ItemID=" & ItemID
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 3
        rs("NewsCollecDate") = Now()
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
    
    If LoginType = 1 And ListNum = 1 Then '�ɼ���¼
        '��¼��վ
        LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
        LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))
        If InStr(LoginResult, LoginFalse) > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڵ�¼��վʱ��������,��ȷ����¼��Ϣ����ȷ�ԣ�</li>"
            ItemNum = ItemNum + 1
            ListNum = 1
            ArticleList = ""
            Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)			
            'Response.Write "   <meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"

            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If
    End If
    
    If FoundErr <> True And ItemEnd <> True Then
        '�����ɼ��б�
        If ArticleList = "" Then  '�����б���
            '�ж��б�����
            '������ҳ����
            If ListPaingType = 0 Then
                If ListNum = 1 Then
                    '�б�����=�б�����ҳ��
                    ListUrl = ListStr
                Else
                    ListEnd = True
                End If
            '���ñ�ǩ
            ElseIf ListPaingType = 1 Then
                '�ж��б�Ϊ1ʱ�������ӵ�ַ
                If ListNum = 1 Then
                    ListUrl = ListStr
                Else
                    If ListPaingNext = "" Or ListPaingNext = "$False$" Then
                        ListEnd = True
                    Else
                        If InStr(ListPaingNext, "{$ID}") > 0 Then
                            ListPaingNext = Replace(ListPaingNext, "{$ID}", "&")
                        End If
                        ListUrl = ListPaingNext
                    End If
                End If
            '��������
            ElseIf ListPaingType = 2 Then
                If ListPaingID1 > ListPaingID2 Then
                    If (ListPaingID1 - ListNum + 1) < ListPaingID2 Or (ListPaingID1 - ListNum + 1) < 0 Then
                        ListEnd = True
                    Else
                        ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1 - ListNum + 1))
                    End If
                Else
                    If (ListPaingID1 + ListNum - 1) > ListPaingID2 Then
                        ListEnd = True
                    Else
                        ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1 + ListNum - 1))
                    End If
                End If
            '�ֶ����
            ElseIf ListPaingType = 3 Then
                ListArray = Split(ListPaingStr3, vbCrLf)
                If (ListNum - 1) > UBound(ListArray) Then
                    ListEnd = True
                Else
                    ListUrl = ListArray(ListNum - 1)
                End If
            End If
            If ListEnd <> True Then
                If CheckUrl(ListStr) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�б���ַ���ԣ�</li>"
                End If
                ArticleList = ListUrl
                If InStr(ListUrl, "{$ID}") > 0 Then
                    ListUrl = Replace(ListUrl, "{$ID}", "&")
                End If
                If FoundErr <> True Then
                    ListCode = GetHttpPage(ListUrl, PE_CLng(WebUrl)) '��ȡ��ҳԴ���� ListCode
                    '����Ϊ���ñ�ǩʱ
                    If ListPaingType = 1 Then
                        ListPaingNext = GetPaing(ListCode, LPsString, LPoString, False, False)
                        If ListPaingNext <> "$False$" Then
                            If ListPaingStr1 <> "" Then
                                ListPaingNext = Replace(ListPaingStr1, "{$ID}", ListPaingNext)
                            Else
                                ListPaingNext = DefiniteUrl(ListPaingNext, ListUrl)
                            End If
                            If InStr(ListPaingNext, "&") > 0 Then
                                ListPaingNext = Replace(ListPaingNext, "&", "{$ID}")
                            End If
                        End If
                    Else
                        ListPaingNext = "$False$"
                    End If
                    
                    If ListCode = "$False$" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�ڻ�ȡ��" & ListUrl & "��ҳԴ��ʱ��������</li>"
                    Else
                        ListCode = GetBody(ListCode, LsString, LoString, False, False) '��ȡ�б��ַ���
                        If ListCode = "$False$" Or ListCode = "" Then
                            FoundErr = True
                            ErrMsg = ErrMsg & "<li>�ڽ�ȡ��" & ListUrl & "�б�ʱ��������</li>"
                        End If
                    End If
                End If
                If FoundErr <> True Then
                    NewsArrayCode = GetArray(ListCode, HsString, HoString, False, False) 'NewsArrayCode=���б�����ȡ���ӵ�ַ
                    If ThumbnailType = 1 Then
                        ThumbnailArrayCode = GetArray(ListCode, ThsString, ThoString, False, False) '����ͼ��ַ
                    End If
                End If

                If NewsArrayCode = "$False$" Or FoundErr = True Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�ڷ�����" & ListUrl & "�����б�ʱ��������</li>"
                    ItemNum = ItemNum + 1
                    ListNum = 1
                    ArticleList = ""
                    '����Html
                    Call GetArrOfCreateHTML
                    Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|")  & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)	
                    'Response.Write "   <meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
                    Call WriteErrMsg(ErrMsg, ComeUrl)

                    Exit Sub
                Else
                    '�ָ��������µ�ַ
                    NewsArray = Split(NewsArrayCode, "$Array$")
                    For Arr_j = 0 To UBound(NewsArray)
                        '�����ӵ�ַҪ���¶�λʱ��
                        If HttpUrlType = 1 Then
                            NewsArray(Arr_j) = Trim(Replace(HttpUrlStr, "{$ID}", NewsArray(Arr_j)))
                        Else
                            '���˿ո񲢽���Ե�ַת��Ϊ���Ե�ַ
                            NewsArray(Arr_j) = Trim(DefiniteUrl(NewsArray(Arr_j), ListUrl))
                        End If
                    Next
                    If PE_CLng(CollecOrder) = 1 Then '����ǵ���ɼ�
                        '�ߵ���ǰ�����˳��
                        For Arr_j = 0 To Fix(UBound(NewsArray) / 2)
                            OrderTemp = NewsArray(Arr_j)
                            NewsArray(Arr_j) = NewsArray(UBound(NewsArray) - Arr_j)
                            NewsArray(UBound(NewsArray) - Arr_j) = OrderTemp
                        Next
                    End If
                    '�б�����ͼ��ַ
                    If ThumbnailType = 1 Then
                        '�ָ��������µ�ַ
                        ThumbnailArray = Split(ThumbnailArrayCode, "$Array$")
                        For Arr_j = 0 To UBound(ThumbnailArray)
                            '���˿ո񲢽���Ե�ַת��Ϊ���Ե�ַ
                            ThumbnailArray(Arr_j) = Trim(DefiniteUrl(ThumbnailArray(Arr_j), ListUrl))
                        Next
                        If PE_CLng(CollecOrder) = 1 Then '����ǵ���ɼ�
                            '�ߵ���ǰ�����˳��
                            For Arr_j = 0 To Fix(UBound(ThumbnailArray) / 2)
                                OrderTemp = ThumbnailArray(Arr_j)
                                ThumbnailArray(Arr_j) = ThumbnailArray(UBound(ThumbnailArray) - Arr_j)
                                ThumbnailArray(UBound(ThumbnailArray) - Arr_j) = OrderTemp
                            Next
                        End If
                        PE_Cache.SetValue "ThumbnailList" & rnd_temp, ThumbnailArray '���ػ���
                    End If
                    PE_Cache.SetValue "ArticleList" & rnd_temp, NewsArray '���ػ���

                    '���¶ϵ��¼
                    sql = "select Timing_Breakpoint from PE_config"
                    Set rs = Server.CreateObject("adodb.recordset")
                    rs.Open sql, Conn, 1, 3
                    rs("Timing_Breakpoint") = " <meta http-equiv=""refresh"" content=0;url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
                    rs.Update
                    rs.Close
                    Set rs = Nothing
                    Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,0)		
                    'Response.Write " <meta http-equiv=""refresh"" content=0;url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
                    Exit Sub
                End If
            Else
                '����Html
                Call GetArrOfCreateHTML

                ErrMsg = ErrMsg & "<br><font color=red>" & ItemName & "</font> ��Ŀ�ɼ�������ɣ�</li>"
                Call WriteSuccessMsg2(ErrMsg)
                Exit Sub
            End If
        Else
            NewsArray = PE_Cache.GetValue("ArticleList" & rnd_temp)
            If ThumbnailType = 1 Then
                ThumbnailArray = PE_Cache.GetValue("ThumbnailList" & rnd_temp)
            End If
        End If
        
        '���ص�����Ϣ
        Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
        Response.Write "    <tr> " & vbCrLf
        Response.Write "      <td height=""22"" colspan=""2"" class=""tdbg"" align=""left"">&nbsp;&nbsp;�ɼ���Ҫһ����ʱ��,�����ĵȴ�,�����վ������ʱ�޷����ʵ��������������,�ɼ��������������󼴿ɻָ���" & vbCrLf
        Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button""  name=""Stop""  value=""ֹͣ�ɼ�""  onCLICK=""location.href='Admin_Collection.asp?Action=StopCollection&rnd_temp=" & rnd_temp & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&CollectionCreateHTML=" & CollectionCreateHTML & "&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&CreateImmediate=" & CreateImmediate & "&UseCreateHTML=" & UseCreateHTML & "&TimingCreate=" & TimingCreate & "'"">" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        '������ʾ�ɼ�ȫ����Ϣ
        Response.Write "    <tr>" & vbCrLf
        Response.Write "      <td height=""22"" colspan=""2"" class=""tdbg"" align=""left"">&nbsp;&nbsp;�������У�" & UBound(ItemIDArray) + 1 & " ����Ŀ,���ڲɼ��� <font color=red>" & ItemNum & "</font> ����Ŀ  <font color=red>" & ItemName & "</font>  �ĵ�   <font color=red>" & ListNum & "</font> ҳ�б�,���б���ɼ�����  <font color=red>" & UBound(NewsArray) + 1 & "</font> ��,�еĵ� <font color=red>" & Arr_i + 1 & "</font> ����" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr>"
        Response.Write "      <td height=""22"" colspan=""2"" class=""tdbg"" align=""left"">&nbsp;&nbsp;�ɼ�ͳ�ƣ��ɹ��ɼ�--" & CollecNewsi & "  ������,ʧ��--" & CollecNewsj & "  ��,ͼƬ--" & ImagesNumAll & "</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "<br>"
        StartTime = Timer()

        If CollecType = 0 Then
            'ִ�к��Ĳɼ�����
            Call StartCollection
            '������������������������������������
            '�ɼ������� ��ǰ�б�ɼ���� or ָ�������
            If CollectionType = 0 And CollectionNum <> "" Then
                If CLng(ItemSucceedNum2) >= CLng(CollectionNum) Then
                    '����Html
                    Call GetArrOfCreateHTML
                    ListNum = ListNum + 1
                    ArticleList = ""
                    ItemSucceedNum2 = 0 'ͳ�����ɼ���Ŀ����0Ϊ��һ���ɼ���Ŀ׼��
                Else
                    Arr_i = Arr_i + 1 '�ƶ�����һ�ɼ�����
                End If
            ElseIf CollectionType = 1 And CollectionNum <> "" Then
                If ListNum > PE_CLng(CollectionNum) Then
                    ArticleList = ""      '�ɼ��б����
                    '����Html
                    Call GetArrOfCreateHTML
                    ItemSucceedNum2 = 0   'ͳ�����ɼ���Ŀ����0Ϊ��һ���ɼ���Ŀ׼��
                Else
                    Arr_i = Arr_i + 1 '�ƶ�����һ�ɼ�����
                End If
            Else
                Arr_i = Arr_i + 1 '�ƶ�����һ�ɼ�����
            End If
            If Arr_i > UBound(NewsArray) Then
                Arr_i = 0
                ListNum = ListNum + 1
                ArticleList = ""      '�ɼ��б����
            End If
        Else
            For Arr_i = 0 To UBound(NewsArray)
                FoundErr = False
                Call StartCollection  'ִ�к��Ĳɼ�����

                '�ɼ������� ��ǰ�б�ɼ���� or ָ�������
                If CollectionType = 0 And CollectionNum <> "" Then
                    If CLng(ItemSucceedNum2) >= CLng(CollectionNum) Then
                        ListNum = ListNum + 1
                        ArticleList = ""
                        '����Html
                        Call GetArrOfCreateHTML
                        ItemSucceedNum2 = 0 'ͳ�����ɼ���Ŀ����0Ϊ��һ���ɼ���Ŀ׼��
                        Exit For
                    End If
                ElseIf PE_CLng(CollectionType) = 1 And CollectionNum <> "" Then
                    If ListNum = PE_CLng(CollectionNum) And Arr_i >= UBound(NewsArray) Then
                        ListNum = ListNum + 1
                        ArticleList = ""      '�ɼ��б����
                        '����Html
                        Call GetArrOfCreateHTML
                        ItemSucceedNum2 = 0   'ͳ�����ɼ���Ŀ����0Ϊ��һ���ɼ���Ŀ׼��
                        Exit For
                    End If
                End If
                If Arr_i >= UBound(NewsArray) Then
                    Arr_i = 0
                    ListNum = ListNum + 1
                    ArticleList = ""      '�ɼ��б����
                    Exit For
                End If
            Next
        End If
    End If

    Response.Write "<br>"
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
    Response.Write "  <tr>"
    Response.Write "   <td height=""22"" align=""left"" class=""tdbg"">"
    Response.Write "&nbsp;&nbsp;����������," & TimeNum & " ������......" & TimeNum & "��������û��Ӧ���� <a href='Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & "'><font color=red>����</font></a> ����<br>"
    Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)
    'Response.Write "   <meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
    Response.Write "&nbsp;&nbsp;ִ��ʱ�䣺" & CStr(FormatNumber((Timer() - StartTime) * 1000, 2)) & " ����"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td height=""22""  class=""tdbg"" align=""left"">&nbsp;&nbsp;�ɼ���Ҫһ����ʱ��,�����ĵȴ�,�����վ������ʱ�޷����ʵ��������������,�ɼ��������������󼴿ɻָ���" & vbCrLf
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button""  name=""Stop""  value=""ֹͣ�ɼ�""  onCLICK=""location.href='Admin_Collection.asp?Action=StopCollection&rnd_temp=" & rnd_temp & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&CollectionCreateHTML=" & CollectionCreateHTML & "&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&CreateImmediate=" & CreateImmediate & "&UseCreateHTML=" & UseCreateHTML & "&TimingCreate=" & TimingCreate & "'"">" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub
'==================================================
'��������StartCollection
'��  �ã���ʼ�ɼ�
'��  ������
'==================================================
Sub StartCollection()

    '������������������������������������
    '����ҳ������ʼ��
    CollecNewsA = CollecNewsA + 1 '�Ѿ��ɼ����������ɹ���ʧ�ܣ�
    DefaultPicUrl = ""   'Ҫ�ɼ��ľ���·��
    ImagesNum = 0        '���βɼ��ɼ�����ͼƬ����
    NewsCode = ""        '�������Ҳ��Դ����
    Title = ""           '����
    Content = ""         '����
    Author = ""          '����
    CopyFrom = ""        '��Դ
    Key = ""             '�ؼ���
    His_Repeat = False   '�Ƿ�ɼ���
    NewsUrl = Trim(NewsArray(Arr_i)) 'Ҫ�ɼ�����������ҳ
    If ThumbnailType = 1 Then
        ThumbnailUrl = Trim(ThumbnailArray(Arr_i))
    End If

    PaingNum = 1               '�������ж��ٷ�ҳ
    UploadFiles = ""           '�ϴ���ͼƬ��ַ
    ErrMsg = ""

    '������������������������������������
    '���ͻ������Ƿ���Ȼ��Ч
    If Response.IsClientConnected Then
        Response.Flush 'ǿ�����Html �������
    Else
        Exit Sub
    End If

    If CollecTest = False Then
        His_Repeat = CheckRepeat(NewsUrl)
    Else
        His_Repeat = False
    End If

    If His_Repeat = True Then
        FoundErr = True
    End If

    '���� ���� ��ȡ����
    If FoundErr <> True Then
        'NewsCode ��ȡ����ҳHtml
        NewsCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl))

        If NewsCode = "$False$" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڻ�ȡ��" & NewsUrl & "����Դ��ʱ��������</li>"
        End If
    End If
    If FoundErr <> True Then
        Title = FpHtmlEnCode(Trim(GetBody(NewsCode, TsString, ToString, False, False))) '��ñ������
        If Title = "$False$" Or Title = "" Or Len(Title) > 200 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڲɼ���" & NewsUrl & "���ű���ʱ��������</li>"
        End If
        If CollecTest = False And IsTitle = True And FoundErr <> True Then
            If PE_CLng(Conn.Execute("Select count(*) From PE_Article Where Title='" & Title & "' And ClassID =" & ClassID)(0)) > 0 Then
                FoundErr = True
            End If
        End If
    End If
    If FoundErr <> True Then
        If CollecType <> 2 Then
            Content = Trim(GetBody(NewsCode, CsString, CoString, False, False)) '������Ĵ���
        End If
        If Content = "$False$" Or Content = "" And CollecType <> 2 Then '�����������Ĳ�������
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڲɼ���" & NewsUrl & "��������ʱ��������</li>"

            If CollecTest = False Then '��Ϊ����ʱ
                'д����ʷ��¼
                sql = "INSERT INTO PE_HistrolyNews(ItemID,ChannelID,ClassID,NewsCollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & Now() & "','" & NewsUrl & "'," & PE_False & ")"
                Conn.Execute (sql)
            End If

        Else
            If CollecType <> 2 Then
                If NewsPaingType = 1 Then '���ŷ�ҳ ���ķ�ҳΪ ���ñ�ǩʱ
                    NewsPaingNext = GetPaing(NewsCode, NPsString, NPoString, False, False) '��ȡ��ҳ��ַ
                    
                    'Ӱ���˲�������ҳ��ҳ��ʱ��ֹ
                    'If Left(NewsPaingNext,1) = "/" then
                        ConversionTrails = NewsUrl
                    'Else
                    '    ConversionTrails = ListStr
                    'End If
                    NewsPaingNext = DefiniteUrl(NewsPaingNext, ConversionTrails) '�����·��ת����·��

                    Dim NewsPaingNextTemp

                    Do While NewsPaingNext <> "$False$" And NewsPaingNext <> ""

                        If NewsPaingNextTemp <> "" Then
                            If FoundInArr(NewsPaingNextTemp, NewsPaingNext, "$$$") = True Then
                                Exit Do
                            Else
                                NewsPaingNextTemp = NewsPaingNextTemp & "$$$" & NewsPaingNext
                            End If

                        Else
                            NewsPaingNextTemp = NewsPaingNext
                        End If
                        
                        If CheckUrl(NewsPaingNext) = False Then
                            Response.Write "<font color=red>����ҳ��ҳ���벻��ȷ,������Ч����ҳ���Ӵ��롣</a>"
                            Exit Do
                        End If

                        NewsPaingNextCode = GetHttpPage(NewsPaingNext, PE_CLng(WebUrl)) '��÷�ҳhtml����

                        If NewsPaingNextCode = "$False$" Or NewsPaingNextCode = "" Then Exit Do
                        ContentTemp = GetBody(NewsPaingNextCode, CsString, CoString, False, False) '��ȡ���Ĵ���

                        If ContentTemp = "$False$" Or ContentTemp = "" Then
                            Exit Do
                        Else
                            PaingNum = PaingNum + 1

                            If PaginationType = 2 Then '��һ�ж���������һ����
                                Content = Content & "<p> </p>[NextPage]<p> </p>" & ContentTemp
                            Else
                                Content = Content & "<p> </p>" & ContentTemp
                            End If
                            '�õ���һ��ҳ���Ӵ���
                            NewsPaingNext = GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False) '��ȡ��ҳ��ַ
                            ''Ӱ���˲�������ҳ��ҳ��ʱ��ֹ
                            'If Left(NewsPaingNext,1) = "/" then
                                ConversionTrails = NewsUrl
                            'Else
                            '    ConversionTrails = ListStr
                            'End If

                            NewsPaingNext = DefiniteUrl(NewsPaingNext, ConversionTrails) '�����·��ת����·��
                        End If

                    Loop

                ElseIf NewsPaingType = 2 Then
                    PageListCode = GetBody(NewsCode, PsString, PoString, False, False) '��ȡ�б�ҳ

                    If PageListCode <> "$False$" Then
                        PageArrayCode = GetArray(PageListCode, PhsString, PhoString, False, False) '��ȡ���ӵ�ַ

                        If PageArrayCode <> "$False$" Then
                            If InStr(PageArrayCode, "$Array$") > 0 Then
                                'ȥ����ַ��ʼ
                                Dim tempk, TempPaingNext
                                PageArray = Split(PageArrayCode, "$Array$") '�ָ�õ���ַ
                                TempPaingNext = ""
                                For tempk = 0 To UBound(PageArray)
                                    If InStr(LCase(TempPaingNext), LCase(PageArray(tempk))) < 1 Then
                                        TempPaingNext = TempPaingNext & "$Array$" & PageArray(tempk)
                                    End If
                                Next
                                TempPaingNext = Right(TempPaingNext, Len(TempPaingNext) - 7)
                                PageArray = Split(TempPaingNext, "$Array$")
                                'ȥ����ַ����

                                For i = 0 To UBound(PageArray)
                                    NewsPaingNextCode = GetHttpPage(DefiniteUrl(PageArray(i), NewsUrl), PE_CLng(WebUrl)) '��÷�ҳhtml����

                                    If NewsPaingNextCode <> "$False$" Or NewsPaingNextCode <> "" Then
                                        ContentTemp = GetBody(NewsPaingNextCode, CsString, CoString, False, False) '��ȡ���Ĵ���

                                        If ContentTemp <> "$False$" Or ContentTemp <> "" Then
                                            PaingNum = PaingNum + 1

                                            If PaginationType = 2 Then '��һ�ж���������һ����
                                                Content = Content & "<p> </p>[NextPage]<p> </p>" & ContentTemp
                                            Else
                                                Content = Content & "<p> </p>" & ContentTemp
                                            End If
                                        End If
                                    End If

                                Next

                            Else
                                NewsPaingNextCode = GetHttpPage(DefiniteUrl(PageArrayCode, NewsUrl), PE_CLng(WebUrl)) '��÷�ҳhtml����

                                If NewsPaingNextCode <> "$False$" Or NewsPaingNextCode <> "" Then
                                    ContentTemp = GetBody(NewsPaingNextCode, CsString, CoString, False, False) '��ȡ���Ĵ���

                                    If ContentTemp <> "$False$" Or ContentTemp <> "" Then
                                        PaingNum = PaingNum + 1

                                        If PaginationType = 2 Then '��һ�ж���������һ����
                                            Content = Content & "<p> </p>[NextPage]<p> </p>" & ContentTemp
                                        Else
                                            Content = Content & "<p> </p>" & ContentTemp
                                        End If
                                    End If
                                End If
                            End If

                        Else
                            Response.Write "<li>�ڻ�ȡ��ҳ�����б�ʱ����</li>"
                        End If

                    Else
                        Response.Write "<li>�ڽ�ȡ��ҳ���뷢������</li>"
                    End If
                End If
            End If
            Call Filters ' ������� ���Ĺ��� ���
        End If
    End If

    If FoundErr <> True Then

        '������������������������������������
        'ʱ��
        If UpDateType = 0 Or UpDateType = "" Then
            UpdateTime = Now()
        ElseIf UpDateType = 1 Then

            If DateType = 0 Then
                UpdateTime = Now()
            Else
                UpdateTime = GetBody(NewsCode, DsString, DoString, False, False)
                UpdateTime = FpHtmlEnCode(UpdateTime)
                UpdateTime = PE_CDate(Trim(Replace(UpdateTime, "&nbsp;", " ")))
            End If

        ElseIf UpDateType = 2 Then
        Else
            UpdateTime = Now()
        End If

        '������������������������������������
        '���߻�ȡ����
        If AuthorType = 1 Then
            Author = GetBody(NewsCode, AsString, AoString, False, False) '��õ�ǰ�����ַ�
        ElseIf AuthorType = 2 Then 'ָ������
            Author = AuthorStr
        Else 'Ϊ0ʱ
            Author = "����"
        End If

        '���߹���
        Author = FpHtmlEnCode(Author)

        If Author = "" Or Author = "$False$" Then
            Author = "����"
        Else

            'ֻ���30���ַ���û���˻�кܳ������֣�
            If Len(Author) > 30 Then
                Author = Left(Author, 30)
            End If
        End If

        '������������������������������������
        '��Դ��ȡ����
        If CopyFromType = 1 Then
            CopyFrom = GetBody(NewsCode, FsString, FoString, False, False)
        ElseIf CopyFromType = 2 Then
            CopyFrom = CopyFromStr
        Else
            CopyFrom = "����"
        End If

        CopyFrom = FpHtmlEnCode(CopyFrom)

        If CopyFrom = "" Or CopyFrom = "$False$" Then
            CopyFrom = "����"
        Else

            If Len(CopyFrom) > 30 Then
                CopyFrom = Left(CopyFrom, 30)
            End If
        End If

        '������������������������������������
        '�ؼ��ֻ�ȡ����
        If KeyType = 0 Then
            Key = Title
            Key = CreateKeyWord(Key, KeyScatterNum)
        ElseIf KeyType = 1 Then
            Key = GetBody(NewsCode, KsString, KoString, False, False)
            Key = FpHtmlEnCode(Key)

            Key = Replace(Key, ",", "|")
            Key = Replace(Key, "&nbsp;", "|")
            Key = Replace(Key, " ", "|")

            Dim arrKey, KeyString, j
            arrKey = Split(Key, "|")
            For j = 0 To UBound(arrKey)
                If arrKey(j) <> "" Then
                    If KeyString = "" Then
                        KeyString = arrKey(j)
                    Else
                        KeyString = KeyString & "|" & arrKey(j)
                    End If
                End If
            Next
            Key = KeyString
            'Key = CreateKeyWord(Key, KeyScatterNum)
            If Len(Key) > 253 Then
                Key = "|" & Left(Key, 253) & "|"
            Else
                Key = "|" & Key & "|"
            End If
        ElseIf KeyType = 2 Then
            Key = KeyStr
            Key = FpHtmlEnCode(Key)

            If Len(Key) > 253 Then
                Key = "|" & Left(Key, 253) & "|"
            Else
                Key = "|" & Key & "|"
            End If
        End If

        If Key = "" Or Key = "$False$" Then
            Key = "|" & Title & "|"
        End If

        '���˷Ƿ��ַ�
        Key = ReplaceBadChar(Key)
        
        '����ɼ��ļ�����·����ַ
        If SaveFlashUrlToFile = True Then
            Content = CollectionFilePath(Content, NewsUrl)
        End If

        '����Զ��ͼƬ
        If CollecTest = False And SaveFiles = True Then
            Content = ReplaceSaveRemoteFile(Content, FilesOverStr, True, FilesPath, NewsUrl)
        Else
            Content = ReplaceSaveRemoteFile(Content, FilesOverStr, False, FilesPath, NewsUrl)
        End If

        
        FilterProperty = Script_Iframe & "|" & Script_Object & "|" & Script_Script & "|" & Script_Class & "|" & Script_Div & "|" & Script_Table & "|" & Script_Tr & "|" & Script_Td & "|" & Script_Span & "|" & Script_Img & "|" & Script_Font & "|" & Script_A & "|" & Script_Html
        Content = FilterScript(Content, FilterProperty) '�ű�����

        If IntroType = 0 Then
        ElseIf IntroType = 1 Then
            Intro = GetBody(NewsCode, IsString, IoString, False, False)
            Intro = Trim(nohtml(Intro))
        ElseIf IntroType = 2 Then
            Intro = nohtml(IntroStr)
        ElseIf IntroType = 3 Then
            Intro = Left(Replace(Replace(Replace(Replace(Trim(nohtml(Content)), vbCrLf, ""), " ", ""), "&nbsp;", ""), "��", ""), IntroNum)
        End If

        If Intro = "$False$" Then
            Intro = ""
        End If
    End If

    If FoundErr <> True Then
        '������������������������������������
        ' ����������ʾ
        If CollecTest = False Then
            Call SaveArticle
            ' ������ʷ��¼
            sql = "INSERT INTO PE_HistrolyNews(ItemID,ChannelID,ClassID,ArticleID,Title,NewsCollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & ArticleID & " ','" & Title & "','" & Now() & "','" & NewsUrl & "'," & PE_True & ")"
            Conn.Execute (sql)
            CollecNewsi = CollecNewsi + 1           '�ɹ��ɼ�����������+1
            ItemSucceedNum = ItemSucceedNum + 1     '��ɲɼ���Ŀ+1
            ItemSucceedNum2 = ItemSucceedNum2 + 1   '�ɹ��ɼ���Ŀ����+1
        End If

        ErrMsg = ErrMsg & "���ű��⣺"

        If CollecTest = False Then
            ErrMsg = ErrMsg & "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & "' target=_blank><font color=red>" & Title & "</font></a>"
        Else
            ErrMsg = ErrMsg & "<font color=red>" & Title & "</font>"
        End If

        ErrMsg = ErrMsg & "<br>"
        ErrMsg = ErrMsg & "�������ߣ�" & Author & "<br>"
        ErrMsg = ErrMsg & "������Դ��" & CopyFrom & "<br>"
        ErrMsg = ErrMsg & "�� �� �֣�"
        If Len(Key) > 4 Then
            ErrMsg = ErrMsg & Mid(Key, 2, Len(Key) - 2) & "<br>"
        Else
             ErrMsg = ErrMsg & Key & "<br>"
        End If
        ErrMsg = ErrMsg & "�ɼ�ҳ�棺<a href=" & NewsUrl & " target=_blank>" & NewsUrl & "</a><br>"
        ErrMsg = ErrMsg & "������Ϣ����ҳ--" & PaingNum & " ҳ,ͼƬ--" & ImagesNum & " ��<br>"

        If Content_view = True And CollecType <> 2 Then
            ErrMsg = ErrMsg & "����Ԥ����"
            ErrMsg = ErrMsg & Left(Content, 250) & "......"
        End If

    Else
        CollecNewsj = CollecNewsj + 1 'ʧ�ܲɼ�����������+1 �����ʷ��¼ His_Repeat�Ƿ���ӹ�����ʷ��¼

        If His_Repeat = True Or IsTitle = True Then
            If CollecType = 0 Then
                TimeNum = 1           '�ɼ���ʱ����
            Else
                TimeNum = 0
            End If

            ErrMsg = ErrMsg & "<li>Ŀ�����ţ�<font color=red>"

            If His_Title = "" Then
                ErrMsg = ErrMsg & NewsUrl '���Ϊ����ʾ��ǰ����
            Else
                ErrMsg = ErrMsg & His_Title
            End If

            ErrMsg = ErrMsg & "</font></a>  �Ѵ���,����ɼ���"
            ErrMsg = ErrMsg & "<li>�ɼ�ʱ�䣺" & His_NewsCollecDate & "</li>"
            ErrMsg = ErrMsg & "<li>������Դ��<a href='" & NewsUrl & "' target=_blank>" & NewsUrl & "</a>"
            ErrMsg = ErrMsg & "<li>��ʾ��Ϣ�������ٴβɼ�,���Ƚ������ŵ���ʷ��¼<font color=red>ɾ��</font></li>"

            If His_Result = True Then
                ErrMsg = ErrMsg & "<li>�Լ������ݿ��е�����ɾ��</li>"
            End If
        End If

        If CollecTest = False And His_Repeat = False And IsTitle = False Then
            sql = "INSERT INTO PE_HistrolyNews(ItemID,ChannelID,ClassID,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & NewsUrl & "'," & PE_False & ")"
            Conn.Execute (sql)
        End If
    End If

    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class='border'>"
    Response.Write "   <tr>"
    Response.Write "      <td height=""22"" colspan=""2"" align=""left"" class=""title"">No:<font color=red>" & CollecNewsA & "</font></td>"
    Response.Write "   </tr>"
    Response.Write "   <tr>"
    Response.Write "      <td colspan=""2"" align=""left"" class=""tdbg"">" & ErrMsg & "</td>"
    Response.Write "   </tr>"
    Response.Write "  <tr>"
    Response.Write "</table>"
    Response.Write "<br>"
End Sub
'=================================================
'��������CheckItem
'��  �ã����������Ŀ�Ƿ�ɲɼ�
'=================================================
Sub CheckItem()
        
    Dim ItemID, ItemName, WebUrl, ListStr, LsString, LoString
    Dim ListPaingType, LPsString, LPoString, ListPaingStr1, ListPaingStr2
    Dim ListPaingID1, ListPaingID2, ListPaingStr3
    Dim HsString, HoString, HttpUrlType, HttpUrlStr
    Dim TsString, ToString, CsString, CoString
    Dim rsItem, sql, FoundErr

    sql = "Select * from PE_Item Where Flag=" & PE_True

    Set rsItem = Server.CreateObject("adodb.recordset")
    rsItem.Open sql, Conn, 1, 1

    If rsItem.EOF And rsItem.BOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����û��,�ɲɼ�����Ŀ,�������ˡ�</li>"
    Else

        Do While Not rsItem.EOF

            ItemID = rsItem("ItemID")
            ItemName = rsItem("ItemName")
            WebUrl = rsItem("WebUrl")
            ListStr = rsItem("ListStr")
            LsString = rsItem("LsString")
            LoString = rsItem("LoString")
            ListPaingType = rsItem("ListPaingType")
            LPsString = rsItem("LPsString")
            LPoString = rsItem("LPoString")
            ListPaingStr1 = rsItem("ListPaingStr1")
            ListPaingStr2 = rsItem("ListPaingStr2")
            ListPaingID1 = rsItem("ListPaingID1")
            ListPaingID2 = rsItem("ListPaingID2")
            ListPaingStr3 = rsItem("ListPaingStr3")

            HsString = rsItem("HsString")
            HoString = rsItem("HoString")
            HttpUrlType = rsItem("HttpUrlType")
            HttpUrlStr = rsItem("HttpUrlStr")

            TsString = rsItem("TsString")
            ToString = rsItem("ToString")
            CsString = rsItem("CsString")
            CoString = rsItem("CoString")
                        
            If LsString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�б�ʼ��ǲ���Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
            End If

            If LoString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�б������ǲ���Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
            End If

            If ListPaingType = 0 Or ListPaingType = 1 Then
                If ListStr = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�б�����ҳ����Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
                End If

                If ListPaingType = 1 Then
                    If LPsString = "" Or LPoString = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>������ҳ��ʼ��������ǲ���Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
                    End If
                End If

                If ListPaingStr1 <> "" And Len(ListPaingStr1) < 15 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>������ҳ�ض������ò���ȷ���޷�����,�뷵����һ���������ã�</li>"
                End If

            ElseIf ListPaingType = 2 Then

                If ListPaingStr2 = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>��������ԭ�ַ�������Ϊ�գ��޷�����,�뷵����һ����������</li>"
                End If

                If IsNumeric(ListPaingID1) = False Or IsNumeric(ListPaingID2) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�������ɵķ�Χֻ�������֣��޷�����,�뷵����һ����������</li>"
                Else

                    If ListPaingID1 = 0 And ListPaingID2 = 0 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�������ɵķ�Χ����ȷ���޷�����,�뷵����һ����������</li>"
                    End If
                End If

            ElseIf ListPaingType = 3 Then

                If ListPaingStr3 = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>������ҳ����Ϊ�գ��޷�����,�뷵����һ����������</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>��ѡ�񷵻���һ������������ҳ����</li>"
            End If

            If HsString = "" Or HoString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>���ӿ�ʼ/������ǲ���Ϊ�գ��޷�����,�뷵����һ����������</li>"
            End If
                        
            If LoginType = 1 Then
                LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
                LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))

                If InStr(LoginResult, LoginFalse) > 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>��¼��վʱ��������,��ȷ�ϵ�¼��Ϣ����ȷ�ԣ�</li>"
                End If
            End If

            If FoundErr <> True Then

                Select Case ListPaingType

                    Case 0, 1
                        ListUrl = ListStr

                    Case 2
                        ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1))

                    Case 3

                        If InStr(ListPaingStr3, vbCrLf) > 0 Then
                            ListUrl = Left(ListPaingStr3, InStr(ListPaingStr3, vbCrLf))
                        Else
                            ListUrl = ListPaingStr3
                        End If

                End Select

            End If

            If FoundErr <> True Then
                ListCode = GetHttpPage(ListUrl, PE_CLng(WebUrl)) '��ȡ��ҳԴ����

                If ListCode <> "$False$" Then
                    ListCode = GetBody(ListCode, LsString, LoString, False, False) '��ȡ�б�ҳ

                    If ListCode <> "$False$" Then
                        NewsArrayCode = GetArray(ListCode, HsString, HoString, False, False) '��ȡ���ӵ�ַ

                        If NewsArrayCode <> "$False$" Then
                            If InStr(NewsArrayCode, "$Array$") > 0 Then
                                NewsArray = Split(NewsArrayCode, "$Array$") '�ָ�õ���ַ

                                If HttpUrlType = 1 Then
                                    NewsUrl = Trim(Replace(HttpUrlStr, "{$ID}", NewsArray(0)))
                                Else
                                    NewsUrl = Trim(DefiniteUrl(NewsArray(0), ListUrl)) 'תΪ����·��
                                End If

                                NewsPaingNextCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl)) '��ȡ��ҳԴ����

                                If NewsPaingType = 1 Then '�������ô����ҳʱ
                                    If NewsPaingStr1 <> "" And Len(NewsPaingStr1) > 15 Then
                                        '��ȡ��ҳ��ַ
                                        ListPaingNext = Replace(NewsPaingStr1, "{$ID}", GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False))
                                    Else
                                        ListPaingNext = GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False) '��ȡ��ҳ��ַ

                                        If ListPaingNext <> "$False$" Then
                                            ListPaingNext = DefiniteUrl(ListPaingNext, NewsUrl) '�����·��ת����·��
                                        End If
                                    End If

                                ElseIf NewsPaingType = 2 Then
                                    PageListCode = GetBody(NewsPaingNextCode, PsString, PoString, False, False) '��ȡ�б�ҳ

                                    If PageListCode <> "$False$" Then
                                        PageArrayCode = GetArray(PageListCode, PhsString, PhoString, False, False) '��ȡ���ӵ�ַ

                                        If PageArrayCode <> "$False$" Then
                                            If InStr(PageArrayCode, "$Array$") > 0 Then
                                                PageArray = Split(PageArrayCode, "$Array$") '�ָ�õ���ַ

                                                For i = 0 To UBound(PageArray)

                                                    If ListPaingNext = "" Then
                                                        ListPaingNext = DefiniteUrl(PageArray(i), NewsUrl) '�����·��ת����·��
                                                    Else
                                                        ListPaingNext = ListPaingNext & "$Array$" & DefiniteUrl(PageArray(i), NewsUrl) '�����·��ת����·��
                                                    End If

                                                    'ȥ����ַ��ʼ
                                                    Dim TempPaingArray, tempj
                                                    TempPaingArray = Split(ListPaingNext, "$Array$")
                                                    ListPaingNext = ""
                                                    For tempj = 0 To UBound(TempPaingArray)
                                                        If InStr(LCase(ListPaingNext), LCase(TempPaingArray(tempj))) < 1 Then
                                                            ListPaingNext = ListPaingNext & "$Array$" & TempPaingArray(tempj)
                                                        End If
                                                    Next
                                                    ListPaingNext = Right(ListPaingNext, Len(ListPaingNext) - 7)
                                                    'ȥ����ַ����

                                                Next

                                            Else
                                                ListPaingNext = DefiniteUrl(PageArrayCode, NewsUrl) '�����·��ת����·��
                                            End If

                                        Else
                                            FoundErr = True
                                            ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�ڻ�ȡ��ҳ�����б�ʱ����</li>"
                                        End If

                                    Else
                                        FoundErr = True
                                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�ڽ�ȡ��ҳ���뷢������</li>"
                                    End If
                                End If

                            Else
                                FoundErr = True
                                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>ֻ����һ����Ч���ӣ���" & NewsArrayCode & "</li>"
                            End If

                        Else
                            FoundErr = True
                            ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�ڻ�ȡ�����б�ʱ����</li>"
                        End If

                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�ڽ�ȡ�б�ʱ��������</li>"
                    End If

                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>�ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������</li>"
                End If
            End If

            If FoundErr <> True Then
                NewsCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl))

                If NewsCode <> "$False$" Then
                    Title = GetBody(NewsCode, TsString, ToString, False, False)
                    Content = GetBody(NewsCode, CsString, CoString, False, False)

                    If Title = "$False$" Or Content = "$False$" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�ڽ�ȡ����/���ĵ�ʱ��������" & NewsUrl & "</li>"
                    End If

                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�ڻ�ȡԴ��ʱ��������" & NewsUrl & "</li>"
                End If
            End If

            If FoundErr = True Then
                Conn.Execute ("update PE_Item set Flag=" & PE_False & " where ItemID=" & ItemID)
            Else
                ErrMsg = "��Ŀ��<font color=red>" & ItemName & "</font>&nbsp;���ɹ�,û�з����κ�����"
            End If

            Response.Write "<br><br>" & vbCrLf
            Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
            Response.Write "  <tr align='center' class='title'><td height='22'><strong>" & ItemName & "��Ŀ�������</strong></td></tr>" & vbCrLf
            Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>" & ErrMsg & "</td></tr>" & vbCrLf
            Response.Write "  <tr align='center' class='tdbg'><td>"
            Response.Write "</td></tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
            FoundErr = False
            ErrMsg = ""

            Response.Flush
            rsItem.MoveNext
        Loop

    End If

    Response.Write "<br><center><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></center>"

    rsItem.Close
    Set rsItem = Nothing

    Call CloseConn
End Sub

'=================================================
'��������CreateItemHtml
'��  �ã��ɼ����Զ�����HTML
'=================================================
Sub CreateItemHtml()
    Dim CollectionCreateHTML, CreateNum, CreateCount, TimingCreateUrl, TimingCreate, ArticleNum
    Dim arrContent, arrContent2

    CollectionCreateHTML = Trim(Request("CollectionCreateHTML"))
    TimingCreate = Trim(Request("TimingCreate"))
    CreateNum = PE_CLng(Trim(Request("CreateNum")))
    CreateCount = UBound(Split(CollectionCreateHTML, "|"))

    If CreateNum <= CreateCount Then
        If InStr(CollectionCreateHTML, "|") > 0 Then
            arrContent = Split(CollectionCreateHTML, "|")
            arrContent2 = Split(arrContent(CreateNum), "$")
            ChannelID = PE_CLng(arrContent2(0))
            ClassID = PE_CLng(arrContent2(1))
            SpecialID = ReplaceBadChar(arrContent2(2))
            ArticleNum = PE_CLng(arrContent2(3))
        Else
            arrContent2 = Split(CollectionCreateHTML, "$")
            ChannelID = PE_CLng(arrContent2(0))
            ClassID = PE_CLng(arrContent2(1))
            SpecialID = ReplaceBadChar(arrContent2(2))
            ArticleNum = PE_CLng(arrContent2(3))
        End If
        TimingCreateUrl = "Admin_CreateArticle.asp?Action=CreateArticle&CreateType=7&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&ArticleNum=" & ArticleNum & "&CreateNum=" & CreateNum & "&ShowBack=No&CollectionCreateHTML=" & CollectionCreateHTML & "&TimingCreate=" & TimingCreate
    Else
        If TimingCreate <> "" Then
            TimingCreateUrl = "Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate
        Else
            Response.Write "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
            Response.Write "<link href='images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
            Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
            Response.Write "  <tr align='center' class='title'><td height='22'><strong>��ϲ����</strong></td></tr>" & vbCrLf
            Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>&nbsp;�ɼ�������Ŀ�������!</td></tr>" & vbCrLf
            Response.Write "  <tr align='center' class='tdbg'><td>"
            Response.Write "</td></tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
            Exit Sub
        End If
    End If

    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function aaa(){window.location.href='" & TimingCreateUrl & "';}" & vbCrLf
    Response.Write "    setTimeout('aaa()',5000);" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub


'**************************************************
'��������GetManagePath
'��  �ã��ɼ���Ŀ����
'��  ����iChannelID ----Ƶ��ID
'����ֵ���ɼ�������
'**************************************************
Function GetManagePath(ByVal iChannelID)
    Dim strPath, sqlPath, rsPath
    iChannelID = PE_CLng(iChannelID)
    strPath = "<IMG SRC='images/img_u.gif' height='12'>�����ڵ�λ�ã��ɼ�" & ItemName & "����&nbsp;&gt;&gt;&nbsp;"

    If iChannelID = 0 Then
        strPath = strPath & "<a href='" & strFileName & "&ChannelID=0'>����Ƶ����Ŀ</a>"
    Else
        sqlPath = "select ChannelID,ChannelName from PE_Channel where ChannelID=" & iChannelID
        Set rsPath = Conn.Execute(sqlPath)

        If rsPath.BOF And rsPath.EOF Then
            strPath = strPath & "�����Ƶ������"
        Else
            strPath = strPath & "<a href='" & strFileName & "&ChannelID=" & rsPath(0) & "'>" & rsPath(1) & "��Ŀ</a>"
        End If

        rsPath.Close
        Set rsPath = Nothing
    End If

    GetManagePath = strPath
End Function
'=================================================
'��������SetCache
'��  �ã�������Ŀ����
'=================================================
Sub SetCache()
    rnd_temp = CStr(rnd_num(5))
    '��ȡ��Ŀ��Ϣ
    sql = "select I.ItemID,I.ItemName,I.ChannelID,I.ClassID,I.SpecialID,I.WebName,I.WebUrl,"
    sql = sql & "I.LoginType,I.LoginUrl,I.LoginPostUrl,I.LoginUser,I.LoginPass,I.LoginFalse,I.ItemDoem,"
    sql = sql & "I.ListStr,I.LsString,I.LoString,I.ListPaingType,I.LPsString,I.LPoString,I.ListPaingStr1,"
    sql = sql & "I.ListPaingStr2,I.ListPaingID1,I.ListPaingID2,I.ListPaingStr3,"
    sql = sql & "I.HsString,I.HoString,I.HttpUrlType,I.HttpUrlStr,"
    sql = sql & "I.TsString,I.ToString,I.CsString,I.CoString,I.AuthorType,"
    sql = sql & "I.AuthorStr,I.AsString,I.AoString,I.CopyFromType,"
    sql = sql & "I.FsString,I.FoString,I.CopyFromStr,I.KeyType,I.KsString,I.KoString,I.KeyStr,I.KeyScatterNum,"
    sql = sql & "I.NewsPaingType,I.NPsString,I.NPoString,I.NewsPaingStr1,I.NewsPaingStr2,"
    sql = sql & "I.PaginationType,I.MaxCharPerPage,I.InfoPoint,"
    sql = sql & "I.OnTop,I.Hot,I.Elite,I.Hits,I.Stars,I.UpdateTime,"
    sql = sql & "I.SkinID,I.TemplateID,I.Script_Iframe,I.Script_Object,I.Script_Script,"
    sql = sql & "I.Script_Class,I.Script_Div,I.Script_Span,I.Script_Img,I.Script_Font,"
    sql = sql & "I.Script_A,I.Script_Html,I.SaveFiles,I.AddWatermark,I.AddThumb,I.CollecOrder,"
    sql = sql & "I.Status,I.CreateImmediate,I.IncludePicYn,I.DefaultPicYn,"
    sql = sql & "I.CollectionNum,I.CollectionType,I.UpDateType,I.DateType,I.DsString,"
    sql = sql & "I.DoString,I.ShowCommentLink,I.Script_Table,I.Script_Tr,I.Script_Td,"
    sql = sql & "I.PsString,I.PoString,I.PhsString,I.PhoString,"
    sql = sql & "I.IsString,I.IoString,I.IntroType,I.IntroStr,I.IntroNum,"
    sql = sql & "I.IsField,I.Field,"
    sql = sql & "I.InfoPurview,I.arrGroupID,I.ChargeType,"
    sql = sql & "I.PitchTime,I.ReadTimes,I.DividePercent,I.SaveFlashUrlToFile,ThumbnailType,ThsString,ThoString,"
    sql = sql & "C.ChannelDir,C.UploadDir,C.UpFileType,C.UseCreateHTML from PE_Item I left join PE_Channel C on I.ChannelID=C.ChannelID where I.ItemID in (" & ItemIDStr & ")"

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 3
    If rs.EOF Then   'û���ҵ�����Ŀ
        ItemEnd = True
        ErrMsg = ErrMsg & "<li>��������,�Ҳ�������Ŀ</li>"
    Else
        PE_Cache.SetValue "Collection" & rnd_temp, rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
    '���ع���
    sql = "Select * from PE_Filters Where Flag=" & PE_True & " order by FilterID ASC"
    Set rs = Conn.Execute(sql)
    If rs.EOF And rs.BOF Then
    Else
        PE_Cache.SetValue "Arr_Filters" & rnd_temp, rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
    '������ʷ��¼
    sql = "select NewsUrl,Title,NewsCollecDate,Result from PE_HistrolyNews"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If Not rs.EOF Then
        PE_Cache.SetValue "Arr_Histrolys" & rnd_temp, rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
End Sub
'=================================================
'��������loadItem
'��  �ã�������Ŀ
'=================================================
Sub loadItem()
    Dim ItemNumTemp
    ItemNumTemp = ItemNum - 1

    ItemID = Arr_Item(0, ItemNumTemp)
    ItemName = Arr_Item(1, ItemNumTemp)
    ChannelID = Arr_Item(2, ItemNumTemp)
    ClassID = Arr_Item(3, ItemNumTemp)      '��Ŀ
    SpecialID = Arr_Item(4, ItemNumTemp)    'ר��
    WebUrl = Arr_Item(6, ItemNumTemp)
    LoginType = Arr_Item(7, ItemNumTemp)
    LoginUrl = Arr_Item(8, ItemNumTemp)
    LoginPostUrl = Arr_Item(9, ItemNumTemp)
    LoginUser = Arr_Item(10, ItemNumTemp)
    LoginPass = Arr_Item(11, ItemNumTemp)
    LoginFalse = Arr_Item(12, ItemNumTemp)
    ListStr = Arr_Item(14, ItemNumTemp)      '�б��ַ
    LsString = Arr_Item(15, ItemNumTemp)     '�б�
    LoString = Arr_Item(16, ItemNumTemp)
    ListPaingType = Arr_Item(17, ItemNumTemp)
    LPsString = Arr_Item(18, ItemNumTemp)
    LPoString = Arr_Item(19, ItemNumTemp)
    ListPaingStr1 = Arr_Item(20, ItemNumTemp)
    ListPaingStr2 = Arr_Item(21, ItemNumTemp)
    ListPaingID1 = Arr_Item(22, ItemNumTemp)
    ListPaingID2 = Arr_Item(23, ItemNumTemp)
    ListPaingStr3 = Arr_Item(24, ItemNumTemp)
    HsString = Arr_Item(25, ItemNumTemp)
    HoString = Arr_Item(26, ItemNumTemp)
    HttpUrlType = Arr_Item(27, ItemNumTemp)
    HttpUrlStr = Arr_Item(28, ItemNumTemp)
    TsString = Arr_Item(29, ItemNumTemp)        '����
    ToString = Arr_Item(30, ItemNumTemp)
    CsString = Arr_Item(31, ItemNumTemp)        '����
    CoString = Arr_Item(32, ItemNumTemp)
    AuthorType = Arr_Item(33, ItemNumTemp)              '����
    AuthorStr = Arr_Item(34, ItemNumTemp)
    AsString = Arr_Item(35, ItemNumTemp)
    AoString = Arr_Item(36, ItemNumTemp)
    CopyFromType = Arr_Item(37, ItemNumTemp)    '��Դ
    FsString = Arr_Item(38, ItemNumTemp)
    FoString = Arr_Item(39, ItemNumTemp)
    CopyFromStr = Arr_Item(40, ItemNumTemp)
    KeyType = Arr_Item(41, ItemNumTemp)         '�ؼ���
    KsString = Arr_Item(42, ItemNumTemp)
    KoString = Arr_Item(43, ItemNumTemp)
    KeyStr = Arr_Item(44, ItemNumTemp)
    KeyScatterNum = Arr_Item(45, ItemNumTemp)
    NewsPaingType = Arr_Item(46, ItemNumTemp)
    NPsString = Arr_Item(47, ItemNumTemp)
    NPoString = Arr_Item(48, ItemNumTemp)
    NewsPaingStr1 = Arr_Item(49, ItemNumTemp)
    NewsPaingStr2 = Arr_Item(50, ItemNumTemp)
    PaginationType = Arr_Item(51, ItemNumTemp)
    MaxCharPerPage = Arr_Item(52, ItemNumTemp)
    InfoPoint = Arr_Item(53, ItemNumTemp)
    OnTop = Arr_Item(54, ItemNumTemp)
    Hot = Arr_Item(55, ItemNumTemp)
    Elite = Arr_Item(56, ItemNumTemp)
    Hits = Arr_Item(57, ItemNumTemp)
    Stars = Arr_Item(58, ItemNumTemp)
    UpdateTime = Arr_Item(59, ItemNumTemp)
    SkinID = Arr_Item(60, ItemNumTemp)
    TemplateID = Arr_Item(61, ItemNumTemp)
    Script_Iframe = Arr_Item(62, ItemNumTemp)
    Script_Object = Arr_Item(63, ItemNumTemp)
    Script_Script = Arr_Item(64, ItemNumTemp)
    Script_Class = Arr_Item(65, ItemNumTemp)
    Script_Div = Arr_Item(66, ItemNumTemp)
    Script_Span = Arr_Item(67, ItemNumTemp)
    Script_Img = Arr_Item(68, ItemNumTemp)
    Script_Font = Arr_Item(69, ItemNumTemp)
    Script_A = Arr_Item(70, ItemNumTemp)
    Script_Html = Arr_Item(71, ItemNumTemp)
    SaveFiles = Arr_Item(72, ItemNumTemp)
    AddWatermark = Arr_Item(73, ItemNumTemp)
    AddThumb = Arr_Item(74, ItemNumTemp)
    CollecOrder = Arr_Item(75, ItemNumTemp)
    Status = Arr_Item(76, ItemNumTemp)
    CreateImmediate = Arr_Item(77, ItemNumTemp)
    IncludePicYn = Arr_Item(78, ItemNumTemp)
    DefaultPicYn = Arr_Item(79, ItemNumTemp)
    CollectionNum = Arr_Item(80, ItemNumTemp)
    CollectionType = Arr_Item(81, ItemNumTemp)
    UpDateType = Arr_Item(82, ItemNumTemp)
    DateType = Arr_Item(83, ItemNumTemp)
    DsString = Arr_Item(84, ItemNumTemp)
    DoString = Arr_Item(85, ItemNumTemp)
    ShowCommentLink = Arr_Item(86, ItemNumTemp)
    Script_Table = Arr_Item(87, ItemNumTemp)
    Script_Tr = Arr_Item(88, ItemNumTemp)
    Script_Td = Arr_Item(89, ItemNumTemp)
    PsString = Arr_Item(90, ItemNumTemp)
    PoString = Arr_Item(91, ItemNumTemp)
    PhsString = Arr_Item(92, ItemNumTemp)
    PhoString = Arr_Item(93, ItemNumTemp)
    IsString = Arr_Item(94, ItemNumTemp)
    IoString = Arr_Item(95, ItemNumTemp)
    IntroType = Arr_Item(96, ItemNumTemp)
    IntroStr = Arr_Item(97, ItemNumTemp)
    IntroNum = Arr_Item(98, ItemNumTemp)
    IsField = Arr_Item(99, ItemNumTemp)
    Field = Arr_Item(100, ItemNumTemp)
    InfoPurview = Arr_Item(101, ItemNumTemp)
    arrGroupID = Arr_Item(102, ItemNumTemp)
    ChargeType = Arr_Item(103, ItemNumTemp)
    PitchTime = Arr_Item(104, ItemNumTemp)
    ReadTimes = Arr_Item(105, ItemNumTemp)
    DividePercent = Arr_Item(106, ItemNumTemp)
    SaveFlashUrlToFile = Arr_Item(107, ItemNumTemp)
    ThumbnailType = Arr_Item(108, ItemNumTemp)
    ThsString = Arr_Item(109, ItemNumTemp)
    ThoString = Arr_Item(110, ItemNumTemp)

    FilesPath = InstallDir & Arr_Item(111, ItemNumTemp) & "/" & Arr_Item(112, ItemNumTemp) & "/" & dirMonth
    FilesOverStr = Arr_Item(113, ItemNumTemp)
    UseCreateHTML = Arr_Item(114, ItemNumTemp)
End Sub
'==================================================
'��������SaveArticle
'��  �ã���������
'��  ������
'==================================================
Sub SaveArticle()

    Dim rsArticle, mrs, arrSpecialID
    Set mrs = Conn.Execute("select max(ArticleID) from PE_Article")

    If IsNull(mrs(0)) Then
        ArticleID = 1
    Else
        ArticleID = mrs(0) + 1
    End If

    Set mrs = Nothing
         
    Set rsArticle = Server.CreateObject("adodb.recordset")
    sql = "select top 1 * from PE_Article"
    rsArticle.Open sql, Conn, 1, 3
    rsArticle.addnew
    rsArticle("ArticleID") = ArticleID
    rsArticle("ChannelID") = ChannelID
    rsArticle("ClassID") = ClassID

    If SpecialID = "" Then
        arrSpecialID = Split("0", ",")
    Else
        arrSpecialID = Split(SpecialID, ",")
    End If
    For i = 0 To UBound(arrSpecialID)
        Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (1," & ArticleID & "," & PE_CLng(arrSpecialID(i)) & ")")
    Next
    rsArticle("Title") = Title
    rsArticle("TitleFontType") = 0
    rsArticle("Intro") = Intro
    If CollecType = 2 Then
        If IsLink = True Then
            rsArticle("Content") = NewsUrl
        Else
            rsArticle("LinkUrl") = NewsUrl
            rsArticle("Content") = ""
        End If
    Else
        rsArticle("Content") = Content
    End If
    rsArticle("Keyword") = Key
    rsArticle("Hits") = Hits
    rsArticle("Author") = Author
    rsArticle("CopyFrom") = CopyFrom

    If IncludePicYn = True Then
        If IncludePic = "" Or IsNull(IncludePic) = True Then
            rsArticle("IncludePic") = 0
        Else
            rsArticle("IncludePic") = IncludePic
        End If
    Else
        rsArticle("IncludePic") = 0
    End If

    rsArticle("OnTop") = OnTop
    rsArticle("Elite") = Elite
    rsArticle("Stars") = Stars
    rsArticle("UpdateTime") = UpdateTime
    rsArticle("CreateTime") = UpdateTime
    rsArticle("PaginationType") = PaginationType
    rsArticle("MaxCharPerPage") = MaxCharPerPage
    rsArticle("SkinID") = SkinID
    rsArticle("TemplateID") = TemplateID

    rsArticle("DefaultPicUrl") = DefaultPicUrl

    If SaveFiles = True Then
        rsArticle("UploadFiles") = UploadFiles
    Else
        rsArticle("UploadFiles") = ""
    End If

    rsArticle("Inputer") = UserName
    rsArticle("Editor") = AdminName
    rsArticle("ShowCommentLink") = ShowCommentLink
    rsArticle("Status") = Status
    rsArticle("Deleted") = False
    rsArticle("PresentExp") = 0
    rsArticle("Receive") = False
    rsArticle("ReceiveUser") = ""
    rsArticle("Received") = ""
    rsArticle("AutoReceiveTime") = 0
    rsArticle("ReceiveType") = 0
    rsArticle("InfoPoint") = InfoPoint
    
    Dim rsField, DefaultValue, ArticleField
    Set rsField = Conn.Execute("select * from PE_Field where (ChannelID=-1 or ChannelID=" & ChannelID & ") and (FieldType=1 or FieldType=2)")
    If Not (rsField.BOF And rsField.EOF) Then
        rsField.MoveFirst

        Do While Not rsField.EOF
            If Trim(rsField("DefaultValue")) = "" Then
                DefaultValue = " "
            End If
            If IsField > 0 Then
                If InStr(Field, "|||") > 0 Then
                    arrField = Split(Field, "|||")

                    For iField = 0 To UBound(arrField)
                        arrField2 = Split(arrField(iField), "{#F}")

                        If rsField("FieldID") = PE_CLng(arrField2(0)) Then
                            FieldType = arrField2(2)
                            FisSting = arrField2(3)
                            FioSting = arrField2(4)
                            FieldStr = arrField2(5)

                            If FieldType = 0 Then
                            ElseIf FieldType = 1 Then
                                ArticleField = GetBody(NewsCode, FisSting, FioSting, False, False)
                                ArticleField = Trim(ArticleField)
                            ElseIf FieldType = 2 Then
                                ArticleField = FieldStr
                            End If

                            Exit For
                        End If

                    Next
                    If rsField("FieldType") = 1 Then
                        If Len(ArticleField) > 253 Then
                            ArticleField = Left(ArticleField, 253)
                        End If
                    End If

                    If Field = "" Or Field = "$False$" Then
                        rsArticle(Trim(rsField("FieldName"))) = DefaultValue
                    Else
                        rsArticle(Trim(rsField("FieldName"))) = ArticleField
                    End If

                Else

                    If InStr(Field, "{#F}") > 0 Then
                        arrField2 = Split(Field, "{#F}")
                        FieldType = PE_CLng(arrField2(2))
                        FisSting = arrField2(3)
                        FioSting = arrField2(4)
                        FieldStr = arrField2(5)

                        If rsField("FieldID") = PE_CLng(arrField2(0)) Then
                            If FieldType = 0 Then
                            ElseIf FieldType = 1 Then
                                ArticleField = GetBody(NewsCode, FisSting, FioSting, False, False)
                                ArticleField = Trim(ArticleField)
                            ElseIf FieldType = 2 Then
                                ArticleField = FieldStr
                            End If
                        End If
                    End If

                    If Field = "" Or Field = "$False$" Then
                        rsArticle(Trim(rsField("FieldName"))) = DefaultValue
                    Else
                        rsArticle(Trim(rsField("FieldName"))) = ArticleField
                    End If
                End If

            Else
                rsArticle(Trim(rsField("FieldName"))) = DefaultValue
            End If

            rsField.MoveNext
        Loop

    End If

    rsField.Close
    Set rsField = Nothing

    rsArticle("InfoPurview") = InfoPurview
    rsArticle("arrGroupID") = arrGroupID
    rsArticle("ChargeType") = ChargeType
    rsArticle("PitchTime") = PitchTime
    rsArticle("ReadTimes") = ReadTimes
    rsArticle("DividePercent") = DividePercent

    rsArticle.Update
    rsArticle.Close
    Set rsArticle = Nothing
    
    If Status = 3 Then
        Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1,ItemChecked=ItemChecked+1 where ChannelID=" & ChannelID & "")
        Conn.Execute ("update PE_Class set ItemCount=ItemCount+1 where ClassID=" & ClassID & "")
    Else
        Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1 where ChannelID=" & ChannelID & "")
    End If

End Sub
'=================================================
'��������HistoryNum
'��  �ã�������Ŀ������
'����    Itemid �����ɼ���Ŀ
'        Result �ɼ���Ŀ�ɹ���ʧ��
'=================================================
Sub HistrolyNum(ByVal ItemID, ByVal Result)
    If IsNumeric(ItemID) = False Then
        Response.Write "�ɼ���Ŀ������"
        Exit Sub
    End If
    sql = "select count(HistrolyNewsID) from PE_HistrolyNews where ItemID=" & ItemID & " and Result=" & Result
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<font color='green'>0</font>"
    Else
        If Result = PE_True Then
            Response.Write "<font color='blue'>" & rs(0) & "</font>"
        Else
            If rs(0) = 0 Then
                Response.Write "<font color='green'>" & rs(0) & "</font>"
            Else
                Response.Write "<font color='red'>" & rs(0) & "</font>"
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub
'=================================================
'��������GetArrOfCreateHTML
'��  �ã�����HTML��ֵ
'=================================================
Sub GetArrOfCreateHTML()
    If CreateImmediate = True And UseCreateHTML <> 0 And ItemSucceedNum2 <> 0 Then
        If CollectionCreateHTML = "" Then
            CollectionCreateHTML = ChannelID & "$" & ClassID & "$" & SpecialID & "$" & ItemSucceedNum2
        Else
            CollectionCreateHTML = CollectionCreateHTML & "|" & ChannelID & "$" & ClassID & "$" & SpecialID & "$" & ItemSucceedNum2
        End If
    End If
End Sub
'=================================================
'��������WriteSuccessMsg2
'��  �ã��ɼ��ɹ���Ϣ
'=================================================
Sub WriteSuccessMsg2(ErrMsg)
    ItemSucceedNum = 0
    ItemSucceedNum2 = 0
    ArticleList = ""
    ItemNum = ItemNum + 1
    ListNum = 1
    Response.Write "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>��ϲ����</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>" & ErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf

    If ItemEnd = False Then
        Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)	
        'Response.Write "<meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
    Else
        PE_Cache.DelCache ("Collection" & Request("rnd_temp"))
        PE_Cache.DelCache ("ArticleList" & Request("rnd_temp"))
        PE_Cache.DelCache ("ThumbnailList" & Request("rnd_temp"))
        PE_Cache.DelCache ("Arr_Filters" & Request("rnd_temp"))
        PE_Cache.DelCache ("Arr_Histrolys" & Request("rnd_temp"))

        Dim arrContent, arrContent2, C_ChannelID, C_ClassID, C_SpecialID, C_ArticleNum
        If CollectionCreateHTML <> "" Then
            Response.Write "<center><FONT style='font-size:12px' color='red'>���Ե�,5���Ӻ�ϵͳ��ʼ���ɲɼ�������¡�</FONT></center>"
            Call Refresh("Admin_Collection.asp?Action=CreateItemHtml&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,5)		
            'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_Collection.asp?Action=CreateItemHtml&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
        Else
            If Timing_AreaCollection = "1" Then
                Response.Write "<center><FONT style='font-size:12px' color='red'>���Ե�,5���Ӻ�ϵͳ��ʼ��������ɼ���</FONT></center>"
                Call Refresh("Admin_AreaCollection.asp?Action=AreaCollectionCreateFile&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,5)				
                'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_AreaCollection.asp?Action=AreaCollectionCreateFile&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
            Else
                If TimingCreate <> "" Then
                    Response.Write "<center><FONT style='font-size:12px' color='red'>���Ե�,5���Ӻ�ϵͳ��ʼ��ʱ���ɡ�</FONT></center>"
                    Call Refresh("Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate,5)					
                    'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate & """>"
                End If
            End If
        End If
    End If
End Sub




'==================================================
'��������UrlEncoding
'��  �ã�ת������
'==================================================
Function UrlEncoding(DataStr)
    On Error Resume Next
    Dim StrReturn, Si, ThisChr, InnerCode, Hight8, Low8
    StrReturn = ""
    For Si = 1 To Len(DataStr)
        ThisChr = Mid(DataStr, Si, 1)
        If Abs(Asc(ThisChr)) < &HFF Then
            StrReturn = StrReturn & ThisChr
        Else
            InnerCode = Asc(ThisChr)
            If InnerCode < 0 Then
                InnerCode = InnerCode + &H10000
            End If
            Hight8 = (InnerCode And &HFF00) \ &HFF
            Low8 = InnerCode And &HFF
            StrReturn = StrReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
        End If
    Next
    UrlEncoding = StrReturn
End Function

'==================================================
'��������ReplaceBadChar2
'��  �ã��滻������ʽ�����ַ�
'��  ����strChar-----Ҫ���˵��ַ�
'����ֵ���滻����ַ�
'==================================================
Function ReplaceBadChar2(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar2 = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "^,(,),*,?,[,],$,+,|,{,}"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "\" & arrBadChar(i))
    Next
    ReplaceBadChar2 = tempChar
End Function

'==================================================
'��������GetPaing
'��  �ã���ȡ��ҳ
'��  ����ConStr   ------Ҫ�ҵ�����
'��  ����StartStr ------������ַͷ��
'��  ����OverStr  ------������ַβ��
'��  ����IncluL   ------�Ƿ�ͳ����ַͷ��
'��  ����IncluR   ------�Ƿ�ͳ����ַβ��
'==================================================
Function GetPaing(ByVal ConStr, StartStr, OverStr, IncluL, IncluR)
    If ConStr = "$False$" Or ConStr = "" Or StartStr = "" Or OverStr = "" Or IsNull(ConStr) = True Or IsNull(StartStr) = True Or IsNull(OverStr) = True Then
        GetPaing = "$False$"
        Exit Function
    End If
    Dim Start, Over, ConTemp, tempStr
    tempStr = LCase(ConStr)
    StartStr = LCase(StartStr)
    OverStr = LCase(OverStr)
    Over = InStr(1, tempStr, OverStr)
    If Over <= 0 Then
        GetPaing = "$False$"
        Exit Function
    Else
        If IncluR = True Then
            Over = Over + Len(OverStr)
        End If
    End If
    tempStr = Mid(tempStr, 1, Over)
    Start = InStrRev(tempStr, StartStr)
    If IncluL = False Then
        Start = Start + Len(StartStr)
    End If
    
    If Start <= 0 Or Start >= Over Then
        GetPaing = "$False$"
        Exit Function
    End If
    ConTemp = Mid(ConStr, Start, Over - Start)
    ConTemp = Trim(ConTemp)
    ConTemp = Replace(ConTemp, " ", "%20")
    ConTemp = Replace(ConTemp, ",", "")
    ConTemp = Replace(ConTemp, "'", "")
    ConTemp = Replace(ConTemp, """", "")
    ConTemp = Replace(ConTemp, ">", "")
    ConTemp = Replace(ConTemp, "<", "")
    ConTemp = Replace(ConTemp, "&nbsp;", "")
    GetPaing = ConTemp
End Function
'==================================================
'��������Filters
'��  �ã�����
'==================================================
Sub Filters()
    If IsNull(Arr_Filters) = True Or IsArray(Arr_Filters) = False Then
        Exit Sub
    End If
    For Filteri = 0 To UBound(Arr_Filters, 2)
        FilterStr = ""
        If Arr_Filters(1, Filteri) = ItemID Or Arr_Filters(1, Filteri) = -1 And Arr_Filters(9, Filteri) = True Then
            If Arr_Filters(3, Filteri) = 1 Then '�������
                If Arr_Filters(4, Filteri) = 1 Then
                    Title = Replace(Title, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
                ElseIf Arr_Filters(4, Filteri) = 2 Then
                    FilterStr = GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Do While FilterStr <> "$False$"
                        Title = Replace(Title, FilterStr, Arr_Filters(8, Filteri))
                        FilterStr = GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Loop
                End If
            ElseIf Arr_Filters(3, Filteri) = 2 Then '���Ĺ���
                If Arr_Filters(4, Filteri) = 1 Then
                    Content = Replace(Content, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
                ElseIf Arr_Filters(4, Filteri) = 2 Then
                    FilterStr = GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Do While FilterStr <> "$False$"
                        Content = Replace(Content, FilterStr, Arr_Filters(8, Filteri))
                        FilterStr = GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Loop
                End If
            End If
        End If
    Next
End Sub

'**************************************************
'��������CreateKeyWord
'��  �ã��ɸ������ַ������ɹؼ���
'��  ����Constr---Ҫ���ɹؼ��ֵ�ԭ�ַ���
'����ֵ�����ɵĹؼ���
'**************************************************
Function CreateKeyWord(ByVal ConStr, ByVal Num)
    If ConStr = "" Or IsNull(ConStr) = True Or ConStr = "$False$" Or IsNumeric(Num) = False Then
        CreateKeyWord = "$False$"
        Exit Function
    End If
    If CLng(Num) < 2 Then
        Num = 2
    End If
    ConStr = Replace(ConStr, Chr(32), "")
    ConStr = Replace(ConStr, Chr(9), "")
    ConStr = Replace(ConStr, "&nbsp;", "")
    ConStr = Replace(ConStr, " ", "")
    ConStr = Replace(ConStr, "(", "")
    ConStr = Replace(ConStr, ")", "")
    ConStr = Replace(ConStr, "<", "")
    ConStr = Replace(ConStr, ">", "")
    Dim i, ConstrTemp
    If Num >= Len(ConStr) Then
        CreateKeyWord = "|" & Left(ConStr, 254) & "|"
        Exit Function
    Else
        For i = 1 To Len(ConStr)
            If i + Num > Len(ConStr) Then
                Exit For
            Else
                ConstrTemp = ConstrTemp & "|" & Mid(ConStr, i, Num)
            End If
        Next
    End If
    If Len(ConstrTemp) < 254 Then
        ConstrTemp = ConstrTemp & "|"
    Else
        ConstrTemp = Left(ConstrTemp, 254) & "|"
    End If
    CreateKeyWord = ConstrTemp
End Function

'==================================================
'��������ReplaceSaveRemoteFile
'��  �ã��滻������Զ���ļ�
'��  ����ConStr ------ Ҫ�滻���ַ���
'��  ����OverStr ----- ������ļ���׺��
'��  ����SaveFiles ------ �Ƿ񱣴��ļ�,False������,True����
'��  ����SaveFilePath- �����ļ���
'��  ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function ReplaceSaveRemoteFile(ConStr, OverStr, SaveFiles, SaveFilePath, TistUrl)
    'On Error Resume Next
    If IsObjInstalled("Microsoft.XMLHTTP") = False Then
        ReplaceSaveRemoteFile = ConStr
        Exit Function
    End If
    If ConStr = "$False$" Or ConStr = "" Then '����Ϊ�ջ���˳�
        ReplaceSaveRemoteFile = "$False$"
        Exit Function
    End If
    OverStr = Replace(Replace(OverStr, "$", ""), " ", "")
    Dim tempStr, TempStr2, tempi, TempArray, TempArray2
    Dim RemoteFileUrl, temptime, SaveFileName, ThumbFileName, StrFileType

    IncludePic = 0

    regEx.Pattern = "<img.+?[^\>]>" '��ѯ���������� <img..>
    Set Matches = regEx.Execute(ConStr)
    For Each Match In Matches
        If tempStr <> "" Then
            tempStr = tempStr & "$Array$" & Match.value '�ۼ�����
        Else
            tempStr = Match.value
        End If
    Next
    If tempStr <> "" Then
        TempArray = Split(tempStr, "$Array$") '�ָ�����
        tempStr = ""
        For tempi = 0 To UBound(TempArray)
            regEx.Pattern = "src\s*=\s*.+?\.(" & OverStr & ")" '��ѯsrc =�ڵ�����
            Set Matches = regEx.Execute(TempArray(tempi))
            For Each Match In Matches
                If tempStr <> "" Then
                    tempStr = tempStr & "$Array$" & Match.value '�ۼӵõ� ���Ӽ�$Array$ �ַ�
                    IncludePic = 2  '��ͼ
                Else
                    tempStr = Match.value
                    IncludePic = 1  'ͼ��
                End If
            Next
        Next
    End If
    If tempStr <> "" Then
        regEx.Pattern = "src\s*=\s*" '���� src =
        tempStr = regEx.Replace(tempStr, "")
    End If

    If ThumbnailType = 1 And ThumbnailUrl <> "" Then '�ɼ��б�����ͼ
        If InStr(tempStr, "$Array$") > 0 Then
            tempStr = ThumbnailUrl & "$Array$" & tempStr
        Else
            tempStr = ThumbnailUrl
        End If
    End If

    If tempStr = "" Or IsNull(tempStr) = True Then '������������û��ͼƬԭֵ����
        ReplaceSaveRemoteFile = ConStr
        Exit Function
    End If

    tempStr = Replace(tempStr, """", "")
    tempStr = Replace(tempStr, "'", "") '����ͼƬ���'"
    If Right(SaveFilePath, 1) = "/" Then '�����ļ����ұ߲�����/
        SaveFilePath = Left(SaveFilePath, Len(SaveFilePath) - 1)
    End If
    If SaveFiles = True Then
        If Not fso.FolderExists(Server.MapPath(SaveFilePath)) Then
            fso.CreateFolder Server.MapPath(SaveFilePath)
            If fso.FolderExists(Server.MapPath(SaveFilePath)) Then
                SaveFiles = True
            Else
                SaveFiles = False
            End If
        End If
    End If
    SaveFilePath = SaveFilePath & "/"
    'ȥ���ظ�ͼƬ��ʼ
    TempArray = Split(tempStr, "$Array$")
    tempStr = ""
    For tempi = 0 To UBound(TempArray)
        If InStr(LCase(tempStr), LCase(TempArray(tempi))) < 1 Then
            tempStr = tempStr & "$Array$" & TempArray(tempi)
        End If
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    TempArray = Split(tempStr, "$Array$")
    'ȥ���ظ�ͼƬ����
    'ת�����ͼƬ��ַ��ʼ
    tempStr = ""
    For tempi = 0 To UBound(TempArray)
        tempStr = tempStr & "$Array$" & DefiniteUrl(Trim(TempArray(tempi)), TistUrl)
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    tempStr = Replace(tempStr, Chr(0), "")

    TempArray2 = Split(tempStr, "$Array$")
    tempStr = ""

    Dim IsThumb

    'ת�����ͼƬ��ַ����
    'ͼƬת��/����
    For tempi = 0 To UBound(TempArray2)
        IsThumb = False
        RemoteFileUrl = TempArray2(tempi)
        If RemoteFileUrl <> "$False$" And RemoteFileUrl <> "" And SaveFiles = True Then '����ͼƬ
            StrFileType = GetFileExt(RemoteFileUrl)

             '����ļ����� �Ƕ�̬���˳�
            If StrFileType = "asp" Or StrFileType = "asa" Or StrFileType = "aspx" Or StrFileType = "php" Or StrFileType = "cer" Or StrFileType = "cdx" Or StrFileType = "exe" Then
                UploadFiles = ""
                ReplaceSaveRemoteFile = ConStr
                Exit Function
            End If
            temptime = GetNumString()
            SaveFileName = temptime & "." & StrFileType      '�����ļ���
            ThumbFileName = temptime & "_S." & StrFileType  '�����ļ���
            
            If SaveRemoteFile(RemoteFileUrl, SaveFilePath & SaveFileName) = True Then
                ConStr = Replace(ConStr, TempArray(tempi), "[InstallDir_ChannelDir]{$UploadDir}/" & dirMonth & "/" & SaveFileName) '�滻ԭ����λ��
                If PhotoObject = 1 Then
                    Dim PE_Thumb
                    Set PE_Thumb = New CreateThumb
                    If tempi = 0 And AddThumb = True Then
                        If PE_Thumb.CreateThumb(SaveFilePath & SaveFileName, SaveFilePath & ThumbFileName, 0, 0) = True Then
                            IsThumb = True
                        End If
                    End If
                    If AddWatermark = True Then
                        Call PE_Thumb.AddWatermark(SaveFilePath & SaveFileName)
                    End If
                    Set PE_Thumb = Nothing
                End If

                If IsThumb = True Then
                    UploadFiles = dirMonth & "/" & ThumbFileName & "|" & dirMonth & "/" & SaveFileName
                Else
                    If UploadFiles = "" Then
                        UploadFiles = dirMonth & "/" & SaveFileName
                    Else
                        UploadFiles = UploadFiles & "|" & dirMonth & "/" & SaveFileName
                    End If
                End If
                If PE_CLng(Trim(Request.Form("IncludePic"))) = 0 Then
                    If tempi > 0 Then
                        IncludePic = 2
                    Else
                        IncludePic = 1
                    End If
                Else
                    IncludePic = PE_CLng(Trim(Request.Form("IncludePic")))
                End If

                If InStr(UploadFiles, "|") = 0 Then
                    DefaultPicUrl = UploadFiles
                    ImagesNum = 1
                Else
                    FilesArray = Split(UploadFiles, "|")
                    DefaultPicUrl = FilesArray(0)
                    ImagesNum = UBound(FilesArray) + 1
                End If
                ImagesNumAll = ImagesNumAll + ImagesNum '�ɼ�ͼƬͳ��
            End If            
        ElseIf RemoteFileUrl <> "$False$" And SaveFiles = False Then '������ͼƬ
            SaveFileName = RemoteFileUrl '�ļ�������ԭ����ͼƬ�ַ�
            ConStr = Replace(ConStr, TempArray(tempi), SaveFileName)
            UploadFiles = ""
            ImagesNum = 0
            If tempi =0 Then
                DefaultPicUrl = RemoteFileUrl
            End If
        End If
    Next
    ReplaceSaveRemoteFile = ConStr
End Function
'==================================================
'��������CollectionFilePath
'��  �ã�����ɼ� ͼƬ ���� �ؼ� �ľ���·�����ı�
'��  ����ConStr ------ Ҫ�滻���ַ���
'��  ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function CollectionFilePath(ConStr, TistUrl)
    On Error Resume Next
    If ConStr = "$False$" Or ConStr = "" Or TistUrl = "" Or TistUrl = "$False$" Then
        CollectionFilePath = ConStr
        Exit Function
    End If
    Dim tempStr, TempStr2, TempStr3, tempi, TempArray, TempArray2, RemoteFileUrl
    Dim f, IsSwftxt, FileName
    IsSwftxt = True
    FileName = InstallDir & "CollectionFilePath.txt"
       
    regEx.Pattern = "(<)(img|param|embed|flash8)(.[^\<]*)(value|src|href)(\s*=)(.[^\<]*)(\.)(gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rmvb)"
    Set Matches = regEx.Execute(ConStr)
    For Each Match In Matches
        If tempStr <> "" Then
            tempStr = tempStr & "$Array$" & Match.value
        Else
            tempStr = Match.value
        End If
    Next
    If tempStr <> "" Then
        regEx.Pattern = "(<)(img|param|embed|flash8)(.[^\<]*)(value|src|href)(\s*=)"
        tempStr = regEx.Replace(tempStr, "")
    End If
        
    If tempStr = "" Or IsNull(tempStr) = True Then
        CollectionFilePath = ConStr
        Exit Function
    End If
    
    tempStr = Replace(tempStr, """", "")
    tempStr = Replace(tempStr, "'", "")

    'ȥ���ظ��ļ���ʼ
    TempArray = Split(tempStr, "$Array$")
    tempStr = ""

    For tempi = 0 To UBound(TempArray)
        If InStr(LCase(tempStr), LCase(TempArray(tempi))) < 1 Then
            tempStr = tempStr & "$Array$" & TempArray(tempi)
        End If
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    TempArray = Split(tempStr, "$Array$")
    'ȥ���ظ��ļ�����
    'ת����Ե�ַ��ʼ
    tempStr = ""
    For tempi = 0 To UBound(TempArray)
        tempStr = tempStr & "$Array$" & DefiniteUrl(TempArray(tempi), TistUrl)
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    tempStr = Replace(tempStr, Chr(0), "")
    TempArray2 = Split(tempStr, "$Array$")
    tempStr = ""
    'ת����Ե�ַ����
    '�滻

    If Not fso.FileExists(Server.MapPath(FileName)) Then
        fso.CreateTextFile (Server.MapPath(FileName))
        If fso.FileExists(Server.MapPath(FileName)) Then
            IsSwftxt = False
        End If
    End If

    If IsSwftxt = True Then
        For tempi = 0 To UBound(TempArray2)
            RemoteFileUrl = TempArray2(tempi)
            regEx.Pattern = TempArray(tempi)
            ConStr = regEx.Replace(ConStr, RemoteFileUrl)
            
            Set f = fso.OpenTextFile(Server.MapPath(FileName), 8, 0)
            If SwfTime = "" Then
                f.Write Chr(13) & Chr(10) & " " & ItemName & Now
                SwfTime = Now
            End If
            f.Write Chr(13) & Chr(10) & RemoteFileUrl
        Next
    End If

    f.Close
    Set f = Nothing
    CollectionFilePath = ConStr
End Function


'**************************************************
'��������CheckRepeat
'��  �ã������ʷ��¼�����ظ�
'��  ����strUrl---��վUrl
'����ֵ��True --- ��
'**************************************************
Function CheckRepeat(strUrl)
    CheckRepeat = False
    If IsArray(Arr_Histrolys) = True Then
        For His_i = 0 To UBound(Arr_Histrolys, 2)
            If Arr_Histrolys(0, His_i) = strUrl Then
                CheckRepeat = True
                His_Title = Arr_Histrolys(1, His_i)
                His_NewsCollecDate = Arr_Histrolys(2, His_i)
                His_Result = Arr_Histrolys(3, His_i)
                Exit For
            End If
        Next
    End If
End Function
'=================================================
'��������rnd_num
'��  �ã�����ָ��λ�õ������
'��  ���������������  ----����
'=================================================
Function rnd_num(rLen)
    Dim ri, rmax, rmin
    rmax = 1
    rmin = 1
    For ri = 1 To rLen + 1
        rmax = rmax * 10
    Next
    rmax = rmax - 1
    For ri = 1 To rLen
        rmin = rmin * 10
    Next
    Randomize
    rnd_num = Int((rnd_num - rmin + 1) * Rnd) + rmin
End Function
%>
