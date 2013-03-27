<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Collection.asp"-->
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
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Collection"   '����Ȩ��


Dim rs, sql, rsItem, rsFilters, rsHistory 'ͨ�ñ���

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>����ɼ�����</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf

Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
Call ShowPageTitle(" �� �� �� �� �� �� ", 10056)
If Trim(Request("Timing_AreaCollection")) = "" Then
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""70"" height=""30""><strong>��������</strong></td>" & vbCrLf
    Response.Write "    <td height=""30""><a href=Admin_AreaCollection.asp?Action=AreaCollectionManage>������ҳ</a> | <a href=""Admin_AreaCollection.asp?Action=AreaCollectionAdd"">�������ɼ���Ŀ</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
End If
Response.Write "</table>"

Select Case Action
    Case "AreaCollectionAdd"
        Call AreaCollectionAdd
    Case "AreaCollectionModify"
        Call AreaCollectionAdd
    Case "AreaCollectionManage"
        Call AreaCollectionManage
    Case "AreaCollectionSave"
        Call AreaCollectionSave
    Case "AreaCollectionDel"
        Call AreaCollectionDel
    Case "AreaCollectionPreviewFile"
        Call AreaCollectionPreviewFile
    Case "AreaCollectionCreateFile"
        Call AreaCollectionCreateFile
    Case Else
        Call AreaCollectionManage
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If

Response.Write "</body></html>"
Call CloseConn


'**************************************************
'��������AreaCollectionAdd
'��  �ã���Ӳɼ�����
'**************************************************
Sub AreaCollectionAdd()
    Dim rsItem, sql
    Dim AreaID, AreaName, AreaFile, AreaIntro, Code, StringReplace, AreaUrl
    Dim LableStart, LableEnd, FilterProperty, UpFileType, AreaPassed
    Dim Script_Property

    FoundErr = False
 
    If Action = "AreaCollectionModify" Then
        AreaID = PE_CLng(Trim(Request("AreaID")))
        If IsNumeric(AreaID) = False Or AreaID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>AreaID,��������</li>"
        End If
        If FoundErr <> True Then
            'ȡ������
            sql = "select * from PE_AreaCollection where AreaID=" & AreaID & " and Type=0"
            Set rsItem = Server.CreateObject("adodb.recordset")
            rsItem.Open sql, Conn, 1, 1
            If rsItem.EOF Then   'û���ҵ�����Ŀ
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���������û���ҵ�����Ŀ��</li>"
            Else
                AreaID = rsItem("AreaID")
                AreaName = rsItem("AreaName")
                AreaFile = rsItem("AreaFile")
                AreaIntro = rsItem("AreaIntro")
                Code = rsItem("Code")
                AreaUrl = rsItem("AreaUrl")
                StringReplace = rsItem("StringReplace")
                LableStart = rsItem("LableStart")
                LableEnd = rsItem("LableEnd")
                FilterProperty = rsItem("FilterProperty")
                UpFileType = rsItem("UpFileType")
                AreaPassed = rsItem("AreaPassed")
            End If
            rsItem.Close
            Set rsItem = Nothing
        End If
        If FoundErr = True Then
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If
    Else
        Code = 0
        FilterProperty = "0|0|0|0|0|0|0|0|0|0|0|0|0"
        UpFileType = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rmvb|html|asp|shtml|jsp|shtml|htm|php|cgi"
    End If

    Response.Write "<script language=""JavaScript"">" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "function setFileFileds(num){    " & vbCrLf
    Response.Write "    for(var i=1,str="""";i<=9;i++){" & vbCrLf
    Response.Write "        eval(""objFiles"" + i +"".style.display='none';"")" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    for(var i=1,str="""";i<=num;i++){" & vbCrLf
    Response.Write "        eval(""objFiles"" + i +"".style.display='';"")" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<br>" & vbCrLf
    Response.Write "<table class=border cellSpacing=1 cellPadding=0 width=""100%"" align=center border=0>" & vbCrLf
    Response.Write "<FORM name=form1 action='Admin_AreaCollection.asp' method=post>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td class=title colSpan=2 height=22>" & vbCrLf
    Response.Write "      <DIV align=center><STRONG>"
    If Action = "AreaCollectionModify" Then
        Response.Write " �� �� �� �� �� �� �� Ŀ "
    Else
        Response.Write " �� �� �� �� �� �� �� Ŀ "
    End If
    Response.Write "</STRONG></DIV></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> �ɼ�������Ŀ���ƣ�&nbsp;</strong></td>" & vbCrLf
    Response.Write "    <td class=""tdbg""><input name=""AreaName"" type=""text"" id=""AreaName"" size=""20"" maxlength=""20"" value=" & AreaName & "> <font color=red> * </font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> �ɼ�������Ŀ��飺&nbsp;</strong></td>" & vbCrLf
    Response.Write "    <td class=""tdbg""> <TEXTAREA NAME='AreaIntro' ROWS='' COLS='' style='width:300px;height:70px'>" & Server.HTMLEncode(AreaIntro) & "</TEXTAREA><font color=red> * </font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> �ļ����ƣ�&nbsp;</strong> </td>" & vbCrLf
    Response.Write "    <td class=""tdbg""><input name=""AreaFile"" type=""text"" id=""AreaFile"" size=""30"" maxlength=""30"" value=" & AreaFile & "> <font color=red> * </font><FONT color='blue'>����: xxxx.html</FONT></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg2'> " & vbCrLf
    Response.Write "    <td height='25' align=""center"" colspan='2' ><strong> ��������</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> ��վURL��&nbsp;</strong></td>" & vbCrLf
    Response.Write "    <td class=""tdbg""><input name=""AreaUrl"" type=""text"" id=""AreaUrl"" size=""50"" maxlength=""100"" value=" & AreaUrl & "> <font color=red> * </font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> ��ҳ�����ʽ��&nbsp;</strong></td>" & vbCrLf
    Response.Write "    <td class=""tdbg"">GB2312��<INPUT TYPE='radio' NAME='Code' value='0' "
    If PE_CLng(Code) = 0 Then Response.Write "checked"
    Response.Write "> UTF-8��<INPUT TYPE='radio' NAME='Code' value='1' "
    If PE_CLng(Code) = 1 Then Response.Write "checked"
    Response.Write "> Big5��<INPUT TYPE='radio' NAME='Code' value='2' "
    If PE_CLng(Code) = 2 Then Response.Write "checked"
    Response.Write "><font color=red> * </font>" & vbCrLf
    Response.Write "     &nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> ��ȡ��ʼ�ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "    <td class=""tdbg""> <TEXTAREA NAME='LableStart' ROWS='' COLS='' style='width:400px;height:70px'>" & Server.HTMLEncode(LableStart) & "</TEXTAREA><font color=red> * </font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> ��ȡ�����ַ���&nbsp;</strong></td>" & vbCrLf
    Response.Write "    <td class=""tdbg""> <TEXTAREA NAME='LableEnd' ROWS='' COLS='' style='width:400px;height:70px'>" & Server.HTMLEncode(LableEnd) & "</TEXTAREA><font color=red> * </font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    Dim arrAreaCode2, arrAreaCode, AreaCode1, AreaCode2, i, ReplaceNum
    arrAreaCode2 = Split(StringReplace, "$$$")
    ReplaceNum = UBound(arrAreaCode2) + 1

    If Action = "AreaCollectionModify" Then
        Response.Write "  <tr class=""tdbg""> " & vbCrLf
        Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> ��ȡ����Ԥ����&nbsp;</strong></td>" & vbCrLf
        Response.Write "    <td class=""tdbg""> <TEXTAREA NAME='preview' ROWS='' COLS='' style='width:500px;height:100px'>" & Server.HTMLEncode(GetBody(GetHttpPage(AreaUrl, PE_CLng(Code)), LableStart, LableEnd, True, True)) & "</TEXTAREA><font color=red> * </font></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
    End If

    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> �ַ��滻��Ŀ����&nbsp;</strong></td>"
    Response.Write "    <td class=""tdbg"">" & vbCrLf
    Response.Write "      <select name=""ReplaceNum"" onChange=""setFileFileds(this.value)"">" & vbCrLf
    Response.Write "         <option value=""0"" " & IsOptionSelected(ReplaceNum, 0) & ">0</option>" & vbCrLf
    Response.Write "         <option value=""1"" " & IsOptionSelected(ReplaceNum, 1) & ">1</option>" & vbCrLf
    Response.Write "         <option value=""2"" " & IsOptionSelected(ReplaceNum, 2) & ">2</option>" & vbCrLf
    Response.Write "         <option value=""3"" " & IsOptionSelected(ReplaceNum, 3) & ">3</option>" & vbCrLf
    Response.Write "         <option value=""4"" " & IsOptionSelected(ReplaceNum, 4) & ">4</option>" & vbCrLf
    Response.Write "         <option value=""5"" " & IsOptionSelected(ReplaceNum, 5) & ">5</option>" & vbCrLf
    Response.Write "         <option value=""6"" " & IsOptionSelected(ReplaceNum, 6) & ">6</option>" & vbCrLf
    Response.Write "         <option value=""7"" " & IsOptionSelected(ReplaceNum, 7) & ">7</option>" & vbCrLf
    Response.Write "         <option value=""8"" " & IsOptionSelected(ReplaceNum, 8) & ">8</option>" & vbCrLf
    Response.Write "         <option value=""9"" " & IsOptionSelected(ReplaceNum, 9) & ">9</option>" & vbCrLf
    Response.Write "      </select>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""></td>" & vbCrLf
    Response.Write "    <td class=""tdbg"">" & vbCrLf
    Response.Write "      <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>" & vbCrLf
    If Action = "AreaCollectionAdd" Then
        For i = 1 To 9
            Response.Write "  <tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "    <td class=""tdbg""  id=""objFiles" & i & """ valign='top' style=""display:'none'"">" & vbCrLf
            Response.Write i
            Response.Write "        ���ַ���<TEXTAREA NAME='ReplaceQuilt" & i & "' ROWS='' COLS='' style='width:250px;height:50px'></TEXTAREA>"
            Response.Write "        �滻Ϊ��<TEXTAREA NAME='ReplaceWith" & i & "' ROWS='' COLS='' style='width:250px;height:50px'></TEXTAREA>"
            Response.Write "    </td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
        Next
    Else
        For i = 0 To UBound(arrAreaCode2)
            arrAreaCode = Split(arrAreaCode2(i), "|||")
            AreaCode1 = arrAreaCode(0)
            AreaCode2 = arrAreaCode(1)

            Response.Write "  <tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "    <td class=""tdbg""  id=""objFiles" & i + 1 & """ valign='top' style=""display:''"">" & vbCrLf
            Response.Write i + 1
            Response.Write "        ���ַ���<TEXTAREA NAME='ReplaceQuilt" & i + 1 & "' ROWS='' COLS='' style='width:250px;height:50px'>" & Server.HTMLEncode(AreaCode1) & "</TEXTAREA>"
            Response.Write "        �滻Ϊ��<TEXTAREA NAME='ReplaceWith" & i + 1 & "' ROWS='' COLS='' style='width:250px;height:50px'>" & Server.HTMLEncode(AreaCode2) & "</TEXTAREA>"
            Response.Write "    </td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
        Next
        ReplaceNum = ReplaceNum + 1
        For i = ReplaceNum To 9
            Response.Write "  <tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "    <td class=""tdbg""  id=""objFiles" & i & """ valign='top' style=""display:'none'"">" & vbCrLf
            Response.Write i
            Response.Write "        ���ַ���<TEXTAREA NAME='ReplaceQuilt" & i & "' ROWS='' COLS='' style='width:250px;height:50px'></TEXTAREA>"
            Response.Write "        �滻Ϊ��<TEXTAREA NAME='ReplaceWith" & i & "' ROWS='' COLS='' style='width:250px;height:50px'></TEXTAREA>"
            Response.Write "    </td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
        Next
    End If
    Response.Write "     </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong> ��ȡ�������ӵĺ�׺����&nbsp;</strong></td>" & vbCrLf
    Response.Write "    <td class=""tdbg""> <input name=""UpFileType"" type=""text"" id=""UpFileType"" size=""50"" maxlength=""50"" value=" & UpFileType & "> <font color=red> * </font> <font color='blue'>ע����|�ָ�</font><br>" & vbCrLf
    Response.Write "  <font color='blue'>˵��:���ɼ����ӵ���Ե�ַת��Ϊ���Ե�ַ,������������Ҫת�����ӵĺ�׺��</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    Script_Property = Split(FilterProperty, "|")

    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width=""150"" class=""tdbg"" align=""right""><strong>�� �� ѡ �&nbsp;</strong></td>"
    Response.Write "    <td height=""22"">"
    Response.Write "      &nbsp;&nbsp;<input name=""Script_Iframe"" type=""checkbox"" id=""Script_Iframe""  value=""1"" "
    If Script_Property(0) = "1" Then Response.Write " checked"
    Response.Write ">Iframe" & vbCrLf
    Response.Write "      <input name=""Script_Object"" type=""checkbox"" id=""Script_Object""  value=""1"" "
    If Script_Property(1) = "1" Then Response.Write " checked"
    Response.Write ">Object" & vbCrLf
    Response.Write "      <input name=""Script_Script"" type=""checkbox"" id=""Script_Script""  value=""1"" "
    If Script_Property(2) = "1" Then Response.Write " checked"
    Response.Write ">Script" & vbCrLf
    Response.Write "      <input name=""Script_Class"" type=""checkbox"" id=""Script_Class""  value=""1"" "
    If Script_Property(3) = "1" Then Response.Write " checked"
    Response.Write ">Style" & vbCrLf
    Response.Write "      <input name=""Script_Div"" type=""checkbox"" id=""Script_Div""  value=""1"" "
    If Script_Property(4) = "1" Then Response.Write " checked"
    Response.Write ">Div" & vbCrLf
    Response.Write "      <input name=""Script_Table"" type=""checkbox"" id=""Script_Table""  value=""1"" "
    If Script_Property(5) = "1" Then Response.Write " checked"
    Response.Write ">Table" & vbCrLf
    Response.Write "      <input name=""Script_Tr"" type=""checkbox"" id=""Script_tr""  value=""1"" "
    If Script_Property(6) = "1" Then Response.Write " checked"
    Response.Write ">Tr" & vbCrLf
    Response.Write "      <input name=""Script_td"" type=""checkbox"" id=""Script_td""  value=""1"" "
    If Script_Property(7) = "1" Then Response.Write " checked"
    Response.Write ">Td" & vbCrLf
    Response.Write "      <br>" & vbCrLf
    Response.Write "      &nbsp;&nbsp;<input name=""Script_Span"" type=""checkbox"" id=""Script_Span""  value=""1"" "
    If Script_Property(8) = "1" Then Response.Write " checked"
    Response.Write ">Span" & vbCrLf
    Response.Write "      &nbsp;&nbsp;<input name=""Script_Img"" type=""checkbox"" id=""Script_Img""  value=""1"" "
    If Script_Property(9) = "1" Then Response.Write " checked"
    Response.Write ">Img&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "      <input name=""Script_Font"" type=""checkbox"" id=""Script_Font""  value=""1"" "
    If Script_Property(10) = "1" Then Response.Write " checked"
    Response.Write ">FONT&nbsp;&nbsp;" & vbCrLf
    Response.Write "      <input name=""Script_A"" type=""checkbox"" id=""Script_A""  value=""1"" "
    If Script_Property(11) = "1" Then Response.Write " checked"
    Response.Write ">A&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "      <input name=""Script_Html"" type=""checkbox"" id=""Script_Html""  value=""1"" "
    If Script_Property(12) = "1" Then Response.Write " checked"
    Response.Write ">Html" & vbCrLf

    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "   <tr class=""tdbg"">" & vbCrLf
    Response.Write "     <td class=""tdbg"" align=""middle"" colSpan=""2"" height=""50"">" & vbCrLf
    Response.Write "       <INPUT id=""AreaID"" type=""hidden"" value=" & AreaID & " name=AreaID>" & vbCrLf
    Response.Write "       <INPUT id=""SaveType"" type=""hidden"" value=""" & Action & """ name=SaveType>" & vbCrLf
    Response.Write "       <INPUT id=""Action"" type=""hidden"" value=""AreaCollectionSave"" name=Action>" & vbCrLf
    Response.Write "       <INPUT type=submit value="" ȷ �� "" name=""Submit"" onclick=""javascript:esave.style.visibility='visible';"">&nbsp;&nbsp;" & vbCrLf
    Response.Write "       <INPUT id=Cancel  type=button value="" ȡ �� "" name=Cancel></td>" & vbCrLf
    Response.Write "   </tr>" & vbCrLf
    Response.Write "   </FORM>" & vbCrLf
    Response.Write "  </table>" & vbCrLf

    Response.Write " <div id=""esave"" style=""position:absolute; top:350px; left:200px; z-index:1;visibility:hidden""> " & vbCrLf
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
'**************************************************
'��������AreaCollectionSave
'��  �ã���������ɼ�����
'**************************************************
Sub AreaCollectionSave()

    Dim rsArea, mrs, SaveType, FoundErr
    Dim AreaID, AreaName, AreaFile, AreaIntro, Code, StringReplace, AreaUrl
    Dim LableStart, LableEnd, UpFileType, FilterProperty, AreaPassed

    Dim Script_Property, ReplaceNum, i, strTemplate
    Dim Script_Iframe, Script_Object, Script_Script, Script_Class
    Dim Script_Div, Script_Span, Script_Img, Script_Font, Script_A, Script_Html
    Dim Script_Table, Script_Tr, Script_Td
    
    FoundErr = False

    AreaID = PE_CLng(Request.Form("AreaID"))
    AreaName = Trim(Request.Form("AreaName"))
    AreaFile = Trim(Request.Form("AreaFile"))
    AreaIntro = Trim(Request.Form("AreaIntro"))
    Code = PE_CLng(Request.Form("Code"))
    StringReplace = Trim(Request.Form("StringReplace"))
    AreaUrl = Request.Form("AreaUrl")
    LableStart = Trim(Request.Form("LableStart"))
    LableEnd = Trim(Request.Form("LableEnd"))
    UpFileType = Trim(Request.Form("UpFileType"))

    Script_Iframe = Trim(Request.Form("Script_Iframe"))
    Script_Object = Trim(Request.Form("Script_Object"))
    Script_Script = Trim(Request.Form("Script_Script"))
    Script_Class = Trim(Request.Form("Script_Class"))
    Script_Div = Trim(Request.Form("Script_Div"))
    Script_Span = Trim(Request.Form("Script_Span"))
    Script_Img = Trim(Request.Form("Script_Img"))
    Script_Font = Trim(Request.Form("Script_Font"))
    Script_A = Trim(Request.Form("Script_A"))
    Script_Html = Trim(Request.Form("Script_Html"))
    Script_Table = Trim(Request.Form("Script_Table"))
    Script_Tr = Trim(Request.Form("Script_Tr"))
    Script_Td = Trim(Request.Form("Script_Td"))

    FilterProperty = Script_Iframe & "|" & Script_Object & "|" & Script_Script & "|" & Script_Class & "|" & Script_Div & "|" & Script_Table & "|" & Script_Tr & "|" & Script_Td & "|" & Script_Span & "|" & Script_Img & "|" & Script_Font & "|" & Script_A & "|" & Script_Html
    
    SaveType = Trim(Request.Form("SaveType"))

    ReplaceNum = PE_CLng(Trim(Request.Form("ReplaceNum")))
    
    If SaveType <> "AreaCollectionModify" Then
        If AreaID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��AreaID��</li>"
        Else
            AreaID = PE_CLng(AreaID)
        End If
    End If
    If AreaName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ɼ���Ŀ���ⲻ��Ϊ��</li>"
    End If
    If AreaFile = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ɼ���ĿJS�ļ�������Ϊ��</li>"
    End If
    If AreaIntro = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ɼ���Ŀ��鲻��Ϊ��</li>"
    End If
    If Code = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������Ŀ�ɼ����벻��Ϊ��</li>"
    End If
    If AreaUrl = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ɼ���վҳ����Ϊ��</li>"
    End If
    If LableStart = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ȡ���뿪ʼ����Ϊ��</li>"
    End If
    If LableEnd = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ȡ�����������Ϊ��</li>"
    End If
    If UpFileType = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ȡ�������ӵĺ�׺������Ϊ��</li>"
    End If
    
    If CheckUrl(AreaUrl) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����,���������ַ���Ǿ���·������վ,����http:// ��ͷʹ�þ���·����</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If GetHttpPage(AreaUrl, PE_CLng(Code)) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ڻ�ȡ:" & AreaUrl & "��ҳԴ��ʱ��������</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    Dim AreaCode
    If FoundErr <> True Then
        AreaCode = GetHttpPage(AreaUrl, PE_CLng(Code)) '����б�Դ����
        If AreaCode <> "" Then
            AreaCode = GetBody(AreaCode, LableStart, LableEnd, True, True) '����б����
            AreaCode = ReplaceStringPath(AreaCode, AreaUrl, UpFileType)
            If AreaCode = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�ڽ�ȡ��������ʱ��������</li>"
            End If
        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڻ�ȡ:" & AreaUrl & "��ҳԴ��ʱ��������</li>"
        End If
    End If
    
    If ReplaceNum <> 0 Then
        For i = 1 To ReplaceNum
            If i <> 1 Then
                StringReplace = StringReplace & "$$$"
            End If
            AreaCode = Replace(AreaCode, Trim(Request("ReplaceQuilt" & i)), Trim(Request("ReplaceWith" & i)))
            StringReplace = StringReplace & Trim(Request("ReplaceQuilt" & i)) & "|||" & Trim(Request("ReplaceWith" & i))
        Next
    End If

    AreaCode = FilterScript(AreaCode, FilterProperty)
    
    strTemplate = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>" & vbCrLf
    strTemplate = strTemplate & "<html>" & vbCrLf
    strTemplate = strTemplate & "<head>" & vbCrLf
    strTemplate = strTemplate & "<title> New Document </title>" & vbCrLf
    strTemplate = strTemplate & "<META http-equiv=Content-Type content=text/html; charset=gb2312><link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
    strTemplate = strTemplate & "</head>" & vbCrLf
    strTemplate = strTemplate & "<body>" & vbCrLf
    strTemplate = strTemplate & vbCrLf & AreaCode & vbCrLf
    strTemplate = strTemplate & "</body>" & vbCrLf
    strTemplate = strTemplate & "</html>" & vbCrLf
    strTemplate = Resumeblank(strTemplate)

    If CreateMultiFolder(InstallDir & "AreaCollection") = False Then '���֧�ִ���Ŀ¼
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ܴ��� AreaCollection �ļ���,�����Ƿ��FSOȨ�޻��Ƿ����վĿ¼д��Ȩ�ޡ�</li>"
        Exit Sub
    End If
    Call WriteToFile(InstallDir & "AreaCollection/" & AreaFile, strTemplate)
    If SaveType = "AreaCollectionAdd" Then
        sql = "SELECT TOP 1 * FROM PE_AreaCollection Where Type=0"
        Set rsArea = Server.CreateObject("adodb.recordset")
        rsArea.Open sql, Conn, 1, 3
        rsArea.addnew
    Else
        sql = "SELECT TOP 1 * FROM PE_AreaCollection where AreaID=" & AreaID & " and Type=0"
        Set rsArea = Server.CreateObject("adodb.recordset")
        rsArea.Open sql, Conn, 1, 3
    End If

    rsArea("AreaName") = AreaName
    rsArea("AreaFile") = AreaFile
    rsArea("AreaIntro") = AreaIntro
    rsArea("Code") = Code
    rsArea("StringReplace") = StringReplace
    rsArea("AreaUrl") = AreaUrl
    rsArea("LableStart") = LableStart
    rsArea("LableEnd") = LableEnd
    rsArea("StringReplace") = StringReplace
    rsArea("FilterProperty") = FilterProperty
    rsArea("UpFileType") = UpFileType
    rsArea("AreaPassed") = True
    rsArea("Type") = 0

    rsArea.Update
    rsArea.Close
    Set rsArea = Nothing

    If SaveType = "AreaCollectionAdd" Then
        Call WriteSuccessMsg("���������Ŀ�ɹ���", "Admin_AreaCollection.asp?Action=AreaCollectionManage")
    Else
        Call WriteSuccessMsg("�޸�������Ŀ�ɹ���", "Admin_AreaCollection.asp?Action=AreaCollectionManage")
    End If

    Call CloseConn
End Sub
'=================================================
'��������AreaCollectionManage()
'��  �ã�����ɼ�����
'=================================================
Sub AreaCollectionManage()

    Dim sql, rs, Action
    Dim rsArea, mrs, SaveType, FoundErr
    Dim AreaID, AreaName, AreaFile, AreaIntro, Code, StringReplace, AreaUrl
    Dim LableStart, LableEnd, FilterProperty, AreaPassed

    Response.Write "<br>" & vbCrLf
    Response.Write "<table class=""border"" border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">" & vbCrLf
    Response.Write "<form name=""myform"" method=""POST"" action=""Admin_AreaCollection.asp"">" & vbCrLf
    Response.Write "  <tr class=""title"" style=""padding: 0px 2px;"">" & vbCrLf
    Response.Write "    <td width=""20"" height=""22"" align=""center""> ID </td>" & vbCrLf
    Response.Write "    <td width=""80"" align=""center""> ����ɼ����� </td>" & vbCrLf
    Response.Write "    <td width=""150"" align=""center""> ����ɼ���� </td>" & vbCrLf
    Response.Write "    <td width=""100"" align=""center"">�����ļ���</td>" & vbCrLf
    Response.Write "    <td width=""200"" align=""center"">���ô���</td> " & vbCrLf
    Response.Write "    <td width=""80"" height=""22"" align=""center""> �� �� �� �� " & vbCrLf
    Response.Write "    <td width=""40"" align=""center""> ��� </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    
    sql = "SELECT * From PE_AreaCollection Where Type=0"

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg' height='50'><td colspan='7' align='center'>ϵͳ����������ɼ���Ŀ��</td></tr></table>"
    Else
        Do While Not rs.EOF
            AreaID = rs("AreaID")
            AreaName = rs("AreaName")
            AreaFile = rs("AreaFile")
            AreaIntro = rs("AreaIntro")
            AreaPassed = rs("AreaPassed")

            Response.Write "<tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""padding: 0px 2px;"">" & vbCrLf
            Response.Write "  <td width=""20"" align=""center"" height=""40"">" & AreaID & " </td>" & vbCrLf
            Response.Write "  <td width=""80"" align=""center"">" & AreaName & "</td> " & vbCrLf
            Response.Write "  <td width=""150"" align=""center"">" & AreaIntro & "</td> " & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">" & AreaFile & "</td> " & vbCrLf
            Response.Write "  <td width=""200"" align=""center""><TEXTAREA NAME='Content' onMouseOver=""this.select()"" style='width:250px;height:50px'>" & "<iframe marginwidth=0 marginheight=0 frameborder=0 src='" & InstallDir & "AreaCollection/" & AreaFile & "'></iframe> " & "</TEXTAREA></td> " & vbCrLf
            Response.Write "  <td width=""80"" align=""center"">"
            Response.Write "    <a href='Admin_AreaCollection.asp?Action=AreaCollectionModify&AreaID=" & AreaID & "' onclick=""javascript:esave.style.visibility='visible';"">�޸�</a>&nbsp;"
            Response.Write "    <a href='Admin_AreaCollection.asp?Action=AreaCollectionDel&AreaID=" & AreaID & "' onClick=""return confirm('ȷ��Ҫɾ������Ŀ��');"">ɾ��</a><br>"
            Response.Write "    <a href='Admin_AreaCollection.asp?Action=AreaCollectionCreateFile&AreaID=" & AreaID & "' onclick=""javascript:esave.style.visibility='visible';"">ˢ��</a>&nbsp;"
            Response.Write "    <a href='Admin_AreaCollection.asp?Action=AreaCollectionPreviewFile&AreaID=" & AreaID & "' >Ԥ��</a>"
            Response.Write "</td> " & vbCrLf
            Response.Write "  <td width=""40"" align=""center"">" & vbCrLf
            If AreaPassed = True Then
                Response.Write "<b>��</b>"
            Else
                Response.Write "<FONT color='red'><b>��</b></FONT>"
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "</tr> " & vbCrLf
            rs.MoveNext
        Loop
        Response.Write "<tr class='tdbg'>" & vbCrLf
        Response.Write "  <td colspan='9' height='32' align='center'>" & vbCrLf
        Response.Write "       <INPUT id=""Action"" type=""hidden"" value=""AreaCollectionCreateFile"" name='Action'>" & vbCrLf
        Response.Write "    <input type=""submit"" value="" ˢ����������ɼ��ļ� "" name=""submit"" onclick=""javascript:esave.style.visibility='visible';"">&nbsp;&nbsp;</td>"
        Response.Write "  </td></tr>" & vbCrLf
        Response.Write "</form>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "<br>" & vbCrLf
        Response.Write "<table border='0' cellpadding='0' cellspacing='1' width='100%' class='border'>" & vbCrLf
        Response.Write " <tr class='tdbg'>" & vbCrLf
        Response.Write "   <td width='120' align='right' class='tdbg5'><strong>����˵����&nbsp;</strong></td>" & vbCrLf
        Response.Write "   <td>����ɼ�,���ǲɼ���վҳ��ĳ���̶�����,��������뱣��Ϊ����ҳ�ṩ��ģ�����,ˢ������ɼ��Ϳ�ʱʱ����.<br><FONT color='red'>��;:</FONT> ���ƴ���վ��¢����Դ,����:�������а�,��Ʊ��Ϣ,Υ�³���,���˽��Ƶ���Щ��Ϣ�ǲ����ṩ�ӿڵ�,ͨ������ɼ��Ϳ�ʱʱ�������±�����</td>" & vbCrLf
        Response.Write " </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "<br>" & vbCrLf
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
    End If
    rs.Close
    Set rs = Nothing
    Call CloseConn
End Sub
'=================================================
'��������AreaCollectionDel()
'��  �ã�����ɼ�ɾ��
'=================================================
Sub AreaCollectionDel()
    Dim AreaID, AreaFile
    AreaID = Trim(Request("AreaID"))
    If AreaID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ����ĿID</li>"
    Else
        AreaID = PE_CLng(AreaID)
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    
    Dim rsArea, FileName
    Set rsArea = Server.CreateObject("adodb.recordset")
    rsArea.Open "select * from PE_AreaCollection where AreaID=" & AreaID & " and Type=0", Conn, 1, 3
    If rsArea.BOF And rsArea.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ"
    Else
        AreaFile = rsArea("AreaFile")
        If FoundErr = False Then
            rsArea.Delete
            rsArea.Update
        End If
    End If
    rsArea.Close
    Set rsArea = Nothing

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If ObjInstalled_FSO = True Then
        FileName = Server.MapPath(InstallDir & "AreaCollection/" & AreaFile)
        If fso.FileExists(FileName) Then
            fso.DeleteFile FileName
        End If
    End If

    Call CloseConn

    Call WriteSuccessMsg("ɾ����" & AreaFile & "��JS�ļ��ɹ���", ComeUrl)
End Sub
'=================================================
'��������AreaCollectionPreviewFile()
'��  �ã�����ɼ�Ԥ��
'=================================================
Sub AreaCollectionPreviewFile()
    Dim AreaID, sqlJs, rsArea
    AreaID = Trim(Request("AreaID"))
    If AreaID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
    Else
        AreaID = PE_CLng(AreaID)
    End If
    sqlJs = "select * from PE_AreaCollection where AreaID=" & AreaID & " and Type=0"
    Set rsArea = Conn.Execute(sqlJs)
    If rsArea.BOF And rsArea.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����JS�ļ���</li>"
        rsArea.Close
        Set rsArea = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>Ԥ���ɼ������ļ�Ч��----" & rsArea("AreaName") & "</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='center'><iframe marginwidth=0 marginheight=0 frameborder=0 width='600' height='350' src='" & InstallDir & "AreaCollection/" & rsArea("AreaFile") & "'></iframe></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='center'><a href='Admin_AreaCollection.asp?Action=AreaCollectionManage'>����</a></td>"
    Response.Write " </tr>"
    Response.Write "  </table>"

    rsArea.Close
    Set rsArea = Nothing
End Sub
'=================================================
'��������AreaCollectionCreateFile()
'��  �ã�����ɼ������ļ�
'=================================================
Sub AreaCollectionCreateFile()

    Dim AreaID, AreaName, AreaFile, AreaIntro, Code, StringReplace, AreaUrl
    Dim LableStart, LableEnd, FilterProperty, UpFileType, AreaPassed
    Dim AreaCode
    Dim sql, Script_Property, rsArea, rsArea2, strTemplate
    Dim Timing_AreaCollection, TimingCreate '��ʱ��������ɼ�
    Dim strSucMsg

    AreaID = PE_Clng(Trim(Request("AreaID")))
    Timing_AreaCollection = Trim(Request("Timing_AreaCollection"))
    TimingCreate = Trim(Request("TimingCreate"))

    If AreaID = 0 Then
        sql = "select * from PE_AreaCollection where AreaPassed=" & PE_True & " and Type=0"
    Else
        sql = "select * from PE_AreaCollection where AreaID=" & AreaID & " and Type=0"
        AreaID = PE_CLng(AreaID)
    End If
   
    Set rsArea = Conn.Execute(sql)
    If rsArea.BOF And rsArea.EOF Then
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���������ļ���</li>"
        rsArea.Close
        Set rsArea = Nothing
        Call WriteErrMsg(ErrMsg, ComeUrl)
        If Timing_AreaCollection = "1" Then
            Response.Write "<center><FONT style='font-size:12px' color='red'>���Ե�,5���Ӻ�ϵͳ��ʼ��ʱ���ɡ�</FONT></center>"
            Call Refresh("Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate,5)			
            'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate & """>"
        End If
        Exit Sub
    Else
        Do While Not rsArea.EOF
            FoundErr = False
            ErrMsg = ""
            AreaID = rsArea("AreaID")
            AreaFile = rsArea("AreaFile")
            Code = rsArea("Code")
            StringReplace = rsArea("StringReplace")
            AreaUrl = rsArea("AreaUrl")
            LableStart = rsArea("LableStart")
            LableEnd = rsArea("LableEnd")
            FilterProperty = rsArea("FilterProperty")
            UpFileType = rsArea("UpFileType")
            AreaPassed = rsArea("AreaPassed")

            If FoundErr <> True Then
                AreaCode = GetHttpPage(AreaUrl, PE_CLng(Code)) '����б�Դ����
                If AreaCode <> "" Then
                    AreaCode = GetBody(AreaCode, LableStart, LableEnd, True, True) '����б����
                    AreaCode = ReplaceStringPath(AreaCode, AreaUrl, UpFileType)
                    If AreaCode = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�ڽ�ȡ��������ʱ��������</li>"
                    End If
                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�ڻ�ȡ:" & AreaUrl & "��ҳԴ��ʱ��������</li>"
                End If
            End If
            
            If FoundErr = True Then
                sql = "update PE_AreaCollection set AreaPassed=" & PE_False & " where AreaID=" & AreaID & " and Type=0"
                Set rsArea2 = Conn.Execute(sql)
                Set rsArea2 = Nothing
            End If
            
            Dim arrAreaCode, arrAreaCode2, i
            If StringReplace <> "" Then
                arrAreaCode = Split(StringReplace, "$$$")
                For i = 0 To UBound(arrAreaCode)
                    arrAreaCode2 = Split(arrAreaCode(i), "|||")
                    AreaCode = Replace(AreaCode, arrAreaCode2(0), arrAreaCode2(1))
                Next
            End If
               
            AreaCode = FilterScript(AreaCode, FilterProperty)

            strTemplate = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>" & vbCrLf
            strTemplate = strTemplate & "<html>" & vbCrLf
            strTemplate = strTemplate & "<head>" & vbCrLf
            strTemplate = strTemplate & "<title> New Document </title>" & vbCrLf
            strTemplate = strTemplate & "<META http-equiv=Content-Type content=text/html; charset=gb2312><link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
            strTemplate = strTemplate & "</head>" & vbCrLf
            strTemplate = strTemplate & "<body>" & vbCrLf
            strTemplate = strTemplate & vbCrLf & AreaCode & vbCrLf
            strTemplate = strTemplate & "</body>" & vbCrLf
            strTemplate = strTemplate & "</html>" & vbCrLf
            strTemplate = Resumeblank(strTemplate)
            If CreateMultiFolder(InstallDir & "AreaCollection") = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���ܴ��� AreaCollection �ļ���,�����Ƿ��FSOȨ�޻��Ƿ����վĿ¼д��Ȩ�ޡ�</li>"
            Else
                Call WriteToFile(InstallDir & "AreaCollection/" & AreaFile, strTemplate)
                strSucMsg = strSucMsg & "<li>���ɡ�" & AreaFile & "�������ļ��ɹ���</li>"
            End If
            rsArea.MoveNext
        Loop
    End If
    rsArea.Close
    Set rsArea = Nothing
    Response.Write "<br>"
    If Timing_AreaCollection = "1" Then
        Response.Write "<center><FONT style='font-size:12px' color='red'>���Ե�,5���Ӻ�ϵͳ��ʼ��ʱ���ɡ�</FONT></center>"
        Call Refresh("Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate,5)
        'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate & """>"
    Else
        Call WriteSuccessMsg(strSucMsg, "Admin_AreaCollection.asp?Action=AreaCollectionManage")
    End If
End Sub

Function IsOptionSelected(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If
End Function




'**************************************************
'��������WriteAreaCollection
'��  �ã���ʾ���������ļ����
'��  ������
'**************************************************
Function WriteAreaCollectionMsg(sErrMsg, AreaBit)
    Dim strMsg
    strMsg = strMsg & "<br>" & vbCrLf
    strMsg = strMsg & "<table cellpadding=2 cellspacing=1 border=0 width=100% class='border' align=center>" & vbCrLf
    If AreaBit = "$True$" Then
        strMsg = strMsg & "  <tr align='center' class='title' ><td><strong> �� ϲ ����</strong></td></tr>" & vbCrLf
    Else
        strMsg = strMsg & "  <tr align='center' class='title' ><td><font color=red><strong>������Ϣ��</strong></font></td></tr>" & vbCrLf
    End If
    strMsg = strMsg & "  <tr class='tdbg'><td height='50' valign='top' ><br>" & sErrMsg & "</td></tr>" & vbCrLf
    strMsg = strMsg & "</table>" & vbCrLf
    WriteAreaCollectionMsg = strMsg
End Function

'**************************************************
'��������Resumeblank
'��  �ã�Html����У��
'����ֵ��У�����Html ����
'**************************************************
Function Resumeblank(ByVal Content)
    If Content = "" Then
        Resumeblank = Content
        Exit Function
    Else
        Content = Trim(Content)
    End If
    Dim strHtml, strHtml2, i, Num, Numtemp, strTemp, arrContent
    strHtml = Replace(Content, "<DIV", "<div")
    strHtml = Replace(strHtml, "</DIV>", "</div>")
    strHtml = Replace(strHtml, "<TABLE", "<table")
    strHtml = Replace(strHtml, "</TABLE>", vbCrLf & "</table>" & vbCrLf)
    strHtml = Replace(strHtml, "<TBODY>", "")
    strHtml = Replace(strHtml, "</TBODY>", "" & vbCrLf)
    strHtml = Replace(strHtml, "<TR", "<tr")
    strHtml = Replace(strHtml, "</TR>", vbCrLf & "</tr>" & vbCrLf)
    strHtml = Replace(strHtml, "<TD", "<td")
    strHtml = Replace(strHtml, "</TD>", "</td>")
    strHtml = Replace(strHtml, "<" & "!--", vbCrLf & "<" & "!--")
    strHtml = Replace(strHtml, "<SELECT", vbCrLf & "<Select")
    strHtml = Replace(strHtml, "</SELECT>", vbCrLf & "</Select>")
    strHtml = Replace(strHtml, "<OPTION", vbCrLf & "  <Option")
    strHtml = Replace(strHtml, "</OPTION>", "</Option>")
    strHtml = Replace(strHtml, "<INPUT", vbCrLf & "  <Input")
    strHtml = Replace(strHtml, "<" & "script", vbCrLf & "<" & "script")
    strHtml = Replace(strHtml, "&amp;", "&")
    strHtml = Replace(strHtml, "{$--", vbCrLf & "<" & "!--$")
    strHtml = Replace(strHtml, "--}", "$--" & ">")
    arrContent = Split(strHtml, vbCrLf)
    For i = 0 To UBound(arrContent)
        Numtemp = False
        If InStr(arrContent(i), "<table") > 0 Then
            Numtemp = True
            If strTemp <> "<table" And strTemp <> "</table>" Then
                Num = Num + 2
            End If
            strTemp = "<table"
        ElseIf InStr(arrContent(i), "<tr") > 0 Then
            Numtemp = True
            If strTemp <> "<tr" And strTemp <> "</tr>" Then
                Num = Num + 2
            End If
            strTemp = "<tr"
        ElseIf InStr(arrContent(i), "<td") > 0 Then
            Numtemp = True
            If strTemp <> "<td" And strTemp <> "</td>" Then
                Num = Num + 2
            End If
            strTemp = "<td"
        ElseIf InStr(arrContent(i), "</table>") > 0 Then
            Numtemp = True
            If strTemp <> "</table>" And strTemp <> "<table" Then
                Num = Num - 2
            End If
            strTemp = "</table>"
        ElseIf InStr(arrContent(i), "</tr>") > 0 Then
            Numtemp = True
            If strTemp <> "</tr>" And strTemp <> "<tr" Then
                Num = Num - 2
            End If
            strTemp = "</tr>"
        ElseIf InStr(arrContent(i), "</td>") > 0 Then
            Numtemp = True
            If strTemp <> "</td>" And strTemp <> "<td" Then
                Num = Num - 2
            End If
            strTemp = "</td>"
        ElseIf InStr(arrContent(i), "<" & "!--") > 0 Then
            Numtemp = True
        End If

        If Num < 0 Then Num = 0
        If Trim(arrContent(i)) <> "" Then
            If i = 0 Then
                strHtml2 = String(Num, " ") & arrContent(i)
            ElseIf Numtemp = True Then
                strHtml2 = strHtml2 & vbCrLf & String(Num, " ") & arrContent(i)
            Else
                strHtml2 = strHtml2 & vbCrLf & arrContent(i)
            End If
        End If
    Next
    Resumeblank = strHtml2
End Function


%>
