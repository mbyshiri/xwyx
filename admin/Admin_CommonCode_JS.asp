<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 1   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim AddType


'������Ա����Ȩ��
If AdminPurview > 1 Then
    PurviewPassed = CheckPurview_Other(AdminPurview_Others, "JsFile_" & ChannelDir)
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
        Call WriteEntry(6, AdminName, "ԽȨ����")
        Response.End
    End If
End If

AddType = PE_CLng(Trim(Request("AddType")))

Response.Write "<html><head><title>JS�������</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle(ChannelName & "����----JS�ļ�����", 10112)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>��������</strong></td><td>"
Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "'>JS�ļ�������ҳ</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=0'>����µ�JS�ļ�����ͨ�б�ʽ��</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=1'>����µ�JS�ļ���ͼƬ�б�ʽ��</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=CreateJS'>ˢ����ĿJS�ļ�</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=CreateJS'>ˢ��ר��JS�ļ�</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    If AddType = 0 Then
        Call Add
    Else
        Call AddPic
    End If
Case "Modify"
    Call Modify
Case "ModifyPic"
    Call ModifyPic
Case "SaveAdd", "SaveModify"
    Call SaveJS_List
Case "SaveAddPic", "SaveModifyPic"
    Call SaveJS_Pic
Case "Preview"
    Call PreviewJS
Case "Del"
    Call DelJS
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Response.Write "<form name='myform' method='post' action='Admin_CreateJS.asp'>"
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='100' height='22'>JS��������</td>"
    Response.Write "    <td>���</td>"
    Response.Write "    <td width='60'>��������</td>"
    Response.Write "    <td>JS�ļ���</td>"
    Response.Write "    <td width='260'>JS���ô���</td>"
    Response.Write "    <td width='100' align='center'>����</td>"
    Response.Write "  </tr>"
    Dim sqlJs, rsJs, JsExists
    sqlJs = "select * from PE_JsFile where ChannelID=" & ChannelID & ""
    Set rsJs = Conn.Execute(sqlJs)
    Do While Not rsJs.EOF
        JsExists = False
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td width='100' align='center'>" & rsJs("JsName") & "</td>"
        Response.Write "    <td>" & PE_HTMLEncode(rsJs("JsReadMe")) & "</td>"
        Response.Write "    <td width='60' align='center'>"
        Select Case rsJs("JsType")
        Case 0
            Response.Write "��ͨ�б�"
        Case 1
            Response.Write "ͼƬ�б�"
        End Select
        Response.Write "    </td>"
        Response.Write "    <td>"
        If ObjInstalled_FSO = True Then
            If fso.FileExists(Server.MapPath(InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName"))) Then
                JsExists = True
            End If
            If JsExists = True Then
                Response.Write rsJs("JsFileName")
            Else
                Response.Write "<font color='red'>" & rsJs("JsFileName") & "</font>"
            End If
        Else
            Response.Write rsJs("JsFileName")
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='260'><textarea name='textarea' cols='36' rows='3'>"
        If rsJs("ContentType") = 1 Then
            Response.Write "<!--#include File=""" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & """-->"
        Else
            Response.Write "<script language='javascript' src='" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & "'></script>"
        End If
        Response.Write "</textarea></td>"
        Response.Write "    <td width='100' align='center'>"
        If rsJs("JsType") = 0 Then
            Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=Modify&ID=" & rsJs("ID") & "'>��������</a>&nbsp;&nbsp;"
        Else
            Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=ModifyPic&ID=" & rsJs("ID") & "'>��������</a>&nbsp;&nbsp;"
        End If
        Response.Write "<a href='Admin_CreateJS.asp?ChannelID=" & ChannelID & "&Action=CreateJs&ID=" & rsJs("ID") & "'>ˢ��</a><br>"
        If JsExists = True Then
            Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=Preview&ID=" & rsJs("ID") & "'>Ԥ��Ч��</a>&nbsp;&nbsp;"
        Else
            Response.Write "<font color='gray'>Ԥ��Ч��</font>&nbsp;&nbsp;"
        End If
        Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=Del&ID=" & rsJs("ID") & "' onclick=""return confirm('���Ҫɾ����JS�ļ���������ļ���ģ����ʹ�ô�JS�ļ�����ע���޸Ĺ���ѽ��');"">ɾ��</a>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        rsJs.MoveNext
    Loop
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellspacing='5' cellpadding='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td align='center'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='CreateAllJs'><input name='ShowBack' type='hidden' id='ShowBack' value='Yes'>"
    Response.Write "    <input type='submit' name='Submit' value='ˢ������JS�ļ�'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<b>˵����</b><br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;��ЩJS������Ϊ�˼ӿ�����ٶ��ر����ɵġ������/�޸�/���/ɾ��" & ChannelShortName & "ʱ��ϵͳ���Զ�ˢ�¸�JS�ļ�����Ҫʱ����Ҳ�����ֶ�ˢ�¡���������µ�JS�ļ�������û�����" & ChannelShortName & "����ʱ�Ϳ����ֶ�ˢ���й�JS�ļ���<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>���ļ���Ϊ��ɫ����ʾ��JS�ļ���û�����ɡ�</font><br>"
    Response.Write "<b>ʹ�÷�����</b><br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�����JS���ô��븴�Ƶ�ҳ���ģ���е����λ�ü��ɡ��ɲμ�ϵͳ�ṩ�ĸ�ҳ�漰ģ�塣"
    Response.Write "</form>"
    rsJs.Close
    Set rsJs = Nothing
End Sub

'******************************************
'��������JsBaseInif
'��  �ã�Js���������Ϣ
'******************************************
Sub JsBaseInif(ByVal JsName, ByVal JsReadme, ByVal ContentType, ByVal JsFileName)
    ContentType = PE_CLng(ContentType)
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>�������ƣ�</strong></td>"
    Response.Write "      <td height='25'><input name='JsName' type='text' id='JsName' value='" & JsName & "' size='49' maxlength='50'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>��飺</strong></td>"
    Response.Write "      <td height='25'><textarea name='JsReadme' cols='40' rows='3' id='JsReadme'>" & JsReadme & "</textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ݴ����ʽ��</strong></td>"
    Response.Write "      <td height='25'> <Input TYPE='radio' Name='ContentType' value='0' " & RadioValue(ContentType, 0) & " onClick=""htmltype.style.display='none';jstype.style.display=''""> JS <Input TYPE='radio' Name='ContentType' value='1' " & RadioValue(ContentType, 1) & "  onClick=""htmltype.style.display='';jstype.style.display='none'""> Html <FONT color='blue'>ע�⣺Ƶ��ѡ������Shtml��ʽʱ��ѡ�ô����������չ��Ϊ.shtml���ļ���ʹ��<br>&lt;!--#include file=""aaaa.html""--&gt;������ָ����������ļ������������������ʹ��JS������û���Ѻá� </font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>�ļ�����</strong></td>"
    Response.Write "      <td height='25'><input name='JsFileName' type='text' id='JsFileName'  value='" & JsFileName & "' size='49' maxlength='50'> <font color='#FF0000'>*</font>"
    Response.Write "       <Span Id='jstype' style=""display:"
    If ContentType = 0 Then
        Response.Write "''"
    Else
        Response.Write "'none'"
    End If
    Response.Write """>"
    Response.Write "<font color='red'>��.jsΪ��չ��</font></Span>"
    Response.Write "<Span Id='htmltype' style=""display:"
    If ContentType = 1 Then
        Response.Write "''"
    Else
        Response.Write "'none'"
    End If
    Response.Write """><font color='red'>��.htmlΪ��չ��</font></Span></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg2'>"
    Response.Write "      <td height='25' colspan='2' align='center' ><strong>��������</strong></td>"
    Response.Write "    </tr>"
End Sub

Sub PreviewJS()
    Dim ID, sqlJs, rsJs
    ID = Trim(Request("ID"))
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
    Else
        ID = PE_CLng(ID)
    End If
    sqlJs = "select * from PE_JsFile where ID=" & ID
    Set rsJs = Conn.Execute(sqlJs)
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����JS�ļ���</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>Ԥ��JS�ļ�Ч��----" & rsJs("JsName") & "</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='center'>"

    If rsJs("ContentType") = 1 Then
        Response.Write "<iframe marginwidth=0 marginheight=0 frameborder=0 name='libin' width='700' height='400' src=" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & "></iframe>"
    Else
        Response.Write "<script language='javascript' src='" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & "'></script>"
    End If

    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='center'>"
    Response.Write "        <a href='javascript:this.location.reload();'>ˢ�±�ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <a href='Admin_ArticleJS.asp?ChannelID=" & ChannelID & "'>������ҳ</a>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"

    rsJs.Close
    Set rsJs = Nothing
End Sub

Sub DelJS()
    Dim ID, sqlJs, rsJs, tJsFileName
    ID = Trim(Request("ID"))
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
    Else
        ID = PE_CLng(ID)
    End If
    sqlJs = "select * from PE_JsFile where ID=" & ID
    Set rsJs = Server.CreateObject("ADODB.Recordset")
    rsJs.Open sqlJs, Conn, 1, 3
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����JS�ļ���</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If
    If ObjInstalled_FSO = True Then
        tJsFileName = Server.MapPath(InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName"))
        If fso.FileExists(tJsFileName) Then
            fso.DeleteFile tJsFileName
        End If
    End If
    rsJs.Delete
    rsJs.Update
    rsJs.Close
    Set rsJs = Nothing
    Call CloseConn
    Response.Redirect "Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID
End Sub

Sub CreateJS(ID)
    Response.Write "<br><iframe id='CreateJS' width='100%' height='100' frameborder='0' src='Admin_CreateJS.asp?ChannelID=" & ChannelID & "&Action=CreateJs&ID=" & ID & "'></iframe>"
End Sub

Function GetClass_Option(CurrentID)
    Dim rsClass, sqlClass, strTemp, tmpDepth, i
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    sqlClass = "Select ClassID,ClassName,ClassType,Depth,NextID from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strTemp = "<option value=''>���������Ŀ</option>"
    Else
        strTemp = ""
        Do While Not rsClass.EOF
            tmpDepth = rsClass(3)
            If rsClass(4) > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            strTemp = strTemp & "<option value='" & rsClass(0) & "'"
            If CurrentID > 0 And rsClass(0) = CurrentID Then
                 strTemp = strTemp & " selected"
            End If
            strTemp = strTemp & ">"
            
            If tmpDepth > 0 Then
                For i = 1 To tmpDepth
                    strTemp = strTemp & "&nbsp;&nbsp;"
                    If i = tmpDepth Then
                        If rsClass(4) > 0 Then
                            strTemp = strTemp & "��&nbsp;"
                        Else
                            strTemp = strTemp & "��&nbsp;"
                        End If
                    Else
                        If arrShowLine(i) = True Then
                            strTemp = strTemp & "��"
                        Else
                            strTemp = strTemp & "&nbsp;"
                        End If
                    End If
                Next
            End If
            strTemp = strTemp & rsClass(1)
            If rsClass(2) = 2 Then
                strTemp = strTemp & "(��)"
            End If
            strTemp = strTemp & "</option>"
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing

    GetClass_Option = strTemp
End Function

%>
