<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim oldKInd

Sub Execute()
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))

    If MaxPerPage <= 0 Then MaxPerPage = 10
       
    FileName = "User_Space.asp?Action=" & Action
    If MaxPerPage > 0 Then strFileName = FileName & "&MaxPerPage=" & MaxPerPage
    If Keyword <> "" Then strFileName = FileName & "&keyword=" & Keyword

    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    If Action = "Template" Then
        Response.Write "function changetemplate(fname){" & vbCrLf
        Response.Write "  var curl = 'User_Space.asp?action=CTemplate&fname=' + fname;" & vbCrLf
        Response.Write "  if(confirm('Ӧ�á�' + fname + '�������Ŀռ䣿')){;" & vbCrLf
        Response.Write "      window.location.href=curl;" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "}" & vbCrLf
    Else
        Response.Write "function CheckInput(){" & vbCrLf
        Response.Write "  if(document.myform.BlogName.value==''){" & vbCrLf
        Response.Write "      alert('���Ʋ���Ϊ�գ�');" & vbCrLf
        Response.Write "      document.myform.BlogName.focus();" & vbCrLf
        Response.Write "      return false;" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  document.myform.Intro.value=editor.HtmlEdit.document.body.innerHTML;" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "function changemode(){" & vbCrLf
        Response.Write "    var dbname=document.myform.addtype.value;" & vbCrLf
        Response.Write "    if(dbname=='2'){" & vbCrLf
        Response.Write "        url.style.display='';" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        url.style.display='none';" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "}" & vbCrLf
    End If
    Response.Write "</script>" & vbCrLf

    If Left(LCase(Action), 5) = "order" Then
        Call SetStat
    Else
        Select Case Action
        Case "Add", "AddRss"
            Call Add
        Case "Modify"
            Call Modify
        Case "SaveAdd", "SaveModify"
            Call SaveBlog
        Case "Del"
            Call Del
        Case "Template"
            Call Template
        Case "CTemplate"
            Call CTemplate
        Case Else
            Call main
        End Select
    End If
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub main()
    Dim rsBlogList, sql, rsuserblog, usespacepass
    If Request("page") <> "" Then
        CurrentPage = CInt(Request("page"))
    Else
        CurrentPage = 1
    End If
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='User_Space.asp'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22' align='center'> "
    Response.Write "            <td width='25'><strong>ID</strong></td>"
    Response.Write "            <td width='120'><strong>����</strong></td>"
    Response.Write "            <td><strong>����</strong></td>"
    Response.Write "            <td width='100'><strong>��������</strong></td>"
    Response.Write "            <td width='70'><strong>��ǰ״̬</strong></td>"
    Response.Write "            <td width='70'><strong>�������</strong></td>"
    Response.Write "            <td width='70'><strong>�� ��</strong></td>"
    Response.Write "          </tr>"

    sql = "select * from PE_Space Where UserID=" & UserID & " order by OrderID"
    Set rsBlogList = Server.CreateObject("ADODB.Recordset")
    rsBlogList.Open sql, Conn, 1, 1
    If rsBlogList.BOF And rsBlogList.EOF Then
        totalPut = 0
        Set rsuserblog = Conn.Execute("Select Blog From PE_User Where UserID=" & UserID)
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br><a href='User_Space.asp?Action=Add'>�������뿪ͨ�ҵľۺϿռ�,</a><br><br></td></tr>"
    Else

        totalPut = rsBlogList.RecordCount
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
                rsBlogList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Dim BlogNum
        Do While Not rsBlogList.EOF
            If rsBlogList("Type") < 2 Then
                If rsBlogList("Passed") = True Then usespacepass = True
                Response.Write "  <tr align='center' bgcolor='#ffbbbb' onmouseout=""this.style.backgroundColor='#ffbbbb'"" onmouseover=""this.style.backgroundColor='#bbbbbb'"">"
            Else
                Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            End If
            Response.Write "<td align='center'>" & rsBlogList("ID") & "</td>"
            Response.Write "<td align='center'>"
            If rsBlogList("Type") < 2 Then Response.Write GetKingName(rsBlogList("ClassID"))
            Response.Write "</td><td>" & rsBlogList("Name") & "</td>"
            Response.Write "<td align='center'>" & FormatDateTime(rsBlogList("BirthDay"), 1) & "</td>"
            Response.Write "<td align='center'>"
            If usespacepass = True Then
                If rsBlogList("Passed") = True Then
                    Response.Write "<font color=""green"">��</font>"
                Else
                    Response.Write "<font color=""red"">��</font>"
                End If
                If rsBlogList("onTop") = True Then
                    Response.Write "&nbsp;<font color=""blue"">��</font>"
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsBlogList("IsElite") = True Then
                    Response.Write "&nbsp;<font color=""green"">��</font>"
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
            Else
                Response.Write "�����..."
            End If
            Response.Write "</td><td align='center'>"
            If rsBlogList("Type") < 2 Then
                Response.Write "<a href='User_Space.asp?Action=Modify&ID=" & rsBlogList("ID") & "'>���ÿռ�</a>&nbsp;"
                Response.Write "&nbsp;</td><td></td>"
            Else
                Response.Write "<a href='User_Space.asp?Action=Modify&ID=" & rsBlogList("ID") & "'>�޸�</a>&nbsp;"
                Response.Write "&nbsp;<a href='User_Space.asp?Action=Del&ID=" & rsBlogList("ID") & "'>ɾ��</a></td>"
                Response.Write "<td><input name='OrderID" & rsBlogList("ID") & "' type='text' id='OrderID" & rsBlogList("ID") & "' value='" & rsBlogList("OrderID") & "' size='4' maxlength='4' style='text-align:center'><input type='submit' name='Submit' value='�޸�' onClick=""document.myform.Action.value='order|" & rsBlogList("ID") & "'""></td>"
            End If
            Response.Write "</tr>"
            BlogNum = BlogNum + 1
            If BlogNum >= MaxPerPage Then Exit Do
            rsBlogList.MoveNext
        Loop
    End If
    rsBlogList.Close
    Set rsBlogList = Nothing
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "���ۺϿռ�", True)
    End If
If usespacepass = True Then
    '��ʾ���ٲ�������
    Response.Write "<br><table align='center'><tr align='center' valign='top'><td width='80'><a href='User_Space.asp?Action=Add'><img src='images/soft_add.gif' border='0'><br>�����Ŀ</a></td>"
    Dim rsItem
    Set rsItem = Conn.Execute("select ID,Name,Type from PE_Space where (Type>=3 and Type<=7) and Passed=" & PE_True & " and UserID=" & UserID & " order by Type desc")
    Do While Not rsItem.EOF
        Select Case rsItem("Type")
        Case 3
            Response.Write "<td width='80'><a href='User_SpaceDiary.asp?Action=Add&ID=" & rsItem("ID") & "'><img src='images/article_add.gif' border='0'><br>����" & rsItem("Name") & "</a></td>"
        Case 4
            Response.Write "<td width='80'><a href='User_SpaceMusic.asp?Action=Add&ID=" & rsItem("ID") & "'><img src='images/article_add.gif' border='0'><br>����" & rsItem("Name") & "</a></td>"
        Case 5
            Response.Write "<td width='80'><a href='User_SpaceBook.asp?Action=Add&ID=" & rsItem("ID") & "'><img src='images/article_add.gif' border='0'><br>����" & rsItem("Name") & "</a></td>"
        Case 6
            Response.Write "<td width='80'><a href='User_SpacePhoto.asp?Action=Add&ID=" & rsItem("ID") & "'><img src='images/photo_add.gif' border='0'><br>����" & rsItem("Name") & "</a></td>"
        Case 7
            Response.Write "<td width='80'><a href='User_SpaceLink.asp?Action=Add&ID=" & rsItem("ID") & "'><img src='images/article_add.gif' border='0'><br>����" & rsItem("Name") & "</a></td>"
        End Select
        rsItem.MoveNext
    Loop
    Set rsItem = Nothing
    Response.Write "</tr></table>"
    'ȡ���û�Ŀ¼��С
    Dim ft, foldersize, strSize, displaysize, usize, D2, spacename
    usize = UserSetting(27)
    spacename = Replace(LCase(UserName & UserID), ".", "")

    If fso.FolderExists(Server.MapPath(InstallDir & "Space/" & spacename & "/")) Then
        Set ft = fso.GetFolder(Server.MapPath(InstallDir & "Space/" & spacename & "/"))
        foldersize = ft.size
        If foldersize = 0 Then foldersize = 1
        displaysize = foldersize / 1048576
        If displaysize < 1 Then
            D2 = 1
        Else
            D2 = Int((displaysize / usize) * 300)
            If D2 > 300 Then D2 = 300
        End If
        strSize = foldersize & "&nbsp;Byte"
        If foldersize > 1024 Then
           foldersize = (foldersize / 1024)
           strSize = FormatNumber(foldersize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;KB"
        End If
        If foldersize > 1024 Then
           foldersize = (foldersize / 1024)
           strSize = FormatNumber(foldersize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;MB"
        End If
        If foldersize > 1024 Then
           foldersize = (SpaceSize / 1024)
           strSize = FormatNumber(foldersize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;GB"
        End If
        Set ft = Nothing
        Response.Write "<br><div align='center'>���Ѿ�ʹ����" & usize & "M�ռ��е�:" & strSize & "<div style=""border: 1px solid #d2d3d9;width: 300px;""><div style=""float: left;width: " & D2 & "px;background:#a2ffa9;""></div></div></div>"
    End If
End If
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td><img src='images/help.gif'>�ۺϿռ䣬�Ǳ�վΪ���ṩ��һ����RSS�����Ķ��������¡������ͼƬ��Ϣ���ܣ�������־��ͼƬ����Ϣ������ۺ���ʾ���ܡ�������ͨ��"
    If usespacepass = True Then
        Response.Write "<a href='User_Space.asp?Action=Add'>������Ŀ</a>"
    Else
        Response.Write "������Ŀ"
    End If
    Response.Write "���������ĸ��˿ռ�ʹ�õ�ģ�顣</td></table>"
End Sub

Sub Add()
If PE_CLng(UserSetting(25)) = 1 Then
    Dim rsBlog, rsBlogClass
    Set rsBlog = Conn.Execute("select top 1 ID,Passed from PE_Space Where Type=1 and UserID=" & UserID)
    If Not (rsBlog.BOF And rsBlog.EOF) Then
        If rsBlog("Passed") = True Then
            Response.Write "<form method='post' action='User_Space.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
            Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
            Response.Write "    <tr class='title'> "
            Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� �� �� �� Ŀ</strong></div></td>"
            Response.Write "    </tr>"
            If Action = "AddRss" Then
                Dim XmlRss, RssDOM, oItem, tetitle, teurl
                teurl = Trim(Request("url"))
                If teurl <> "" Then
                    On Error Resume Next
                    Set XmlRss = Server.CreateObject("MSXML2.ServerXMLHTTP")
                    XmlRss.SetTimeouts 5000, 5000, 120000, 60000
                    XmlRss.Open "GET", teurl, False
                    XmlRss.Send
                    If Err.Number <> 0 Then
                        Err.Clear
                        Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Դ״̬��</strong><font color='#FF0000'>��Դ��ַ�����ڻ��޷�����!</font></td></tr>"
                    Else
                        If XmlRss.Readystate <> 4 Or Trim(XmlRss.responseText & "") = "" Then
                            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Դ״̬��</strong><font color='#FF0000'>��Դ��ַ�����ڻ��޷�����!</font></td></tr>"
                        Else
                            Set RssDOM = Server.CreateObject("microsoft.XMLDOM")
                            RssDOM.async = False
                            RssDOM.Load (XmlRss.responseXML)
                            If RssDOM.Readystate = 4 Then
                                Dim RSSVersion, rootNode
                                Set rootNode = RssDOM.documentElement
                                Select Case rootNode.NodeName
                                Case "rss"
                                    RSSVersion = "RSS" & rootNode.getAttribute("version")
                                    If rootNode.getAttribute("version") = "2.0" Then
                                        Set oItem = RssDOM.getElementsByTagName("channel")
                                        tetitle = oItem(0).selectSingleNode("title").text
                                    End If
                                Case "rdf:RDF"
                                    RSSVersion = "RSS1.0"
                                    Set oItem = RssDOM.getElementsByTagName("channel")
                                    tetitle = oItem(0).selectSingleNode("title").text
                                Case "feed"
                                    RSSVersion = "ATOM"
                                    Set oItem = RssDOM.getElementsByTagName("feed")
                                    tetitle = oItem(0).selectSingleNode("title").text
                                End Select
                            Else
                                Response.Write "<tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Դ״̬��</strong><font color='#FF0000'>�����ַ������Ч��RSS����Դ!</font></td></tr>"
                            End If
                        Set RssDOM = Nothing
                        End If
                    End If
                    Set XmlRss = Nothing
                End If
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Ŀ���ͣ�</strong><select name='addtype' onChange=""changemode()""><option value=2 Selected>�ⲿRSS����</option><option value=3>�ҵ���־</option><option value=4>�ҵ�����</option><option value=5>�ҵ�ͼ��</option><option value=6>�ҵ�ͼƬ</option><option value=7>�ҵ�����</option></select> <font color='#FF0000'>* ����Ŀʹ�õĿռ�ģ�� <a href='space_detal.html' target='_blank'>[�鿴��ϸ˵��]</a></font></td></tr>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Ŀ���ƣ�</strong><input name='BlogName' type='text' size='45' maxlength='40' value='" & tetitle & "'> <font color='#FF0000'>*"
                If RSSVersion <> "" Then Response.Write "�������һ����" & RSSVersion & "����ʽ������Դ"
                Response.Write "</font></td></tr>"
                Response.Write "<tbody id='url' style='display:'>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Դ��ַ��</strong><input name='LinkUrl' type='text' size='45' maxlength='100' value='" & teurl & "'> <font color='#FF0000'>* ֧��RSS1.0 RSS2.0 ATOM��ʽ</font></td></tr></tbody>"
            Else
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Ŀ���ͣ�</strong><select name='addtype' onChange=""changemode()""><option value=2>�ⲿRSS����</option><option value=3>�ҵ���־</option><option value=4>�ҵ�����</option><option value=5>�ҵ�ͼ��</option><option value=6>�ҵ�ͼƬ</option><option value=7>�ҵ�����</option></select> <font color='#FF0000'>* ����Ŀʹ�õĿռ�ģ�� <a href='space_detal.html' target='_blank'>[�鿴��ϸ˵��]</a></font></td></tr>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Ŀ���ƣ�</strong><input name='BlogName' type='text' size='45' maxlength='40'> <font color='#FF0000'>*</font></td></tr>"
                Response.Write "<tbody id='url' style='display:'>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Դ��ַ��</strong><input name='LinkUrl' type='text' size='45' maxlength='100' value='http://'> <font color='#FF0000'>* ֧��RSS1.0 RSS2.0 ATOM��ʽ</font></td></tr></tbody>"
            End If
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��ʾ������</strong><input name='ListNum' type='text' size='5' maxlength='3' value='10'> <font color='#FF0000'>* ��ҳ���ҳ��ʾ����</font></td></tr>"
            Response.Write "  <tr class='tdbg'> "
            Response.Write "    <td colspan='2'>&nbsp;<strong>��Ŀ��ҳ��ʾ����</strong>��<br>"
            Response.Write "      <textarea name='Intro' id='Intro' cols='72' rows='9' style='display:none'></textarea>"
            Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='300' ></iframe>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            Response.Write "  <tr>"
            Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
            Response.Write "    <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
            Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Space.asp?Action=Manage';"" style='cursor:hand;'></td>"
            Response.Write "  </tr>"
            Response.Write "</table></form>"
        Else
            Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
            Response.Write "    <tr class='title'> "
            Response.Write "      <td height='22'> <div align='center'><strong>�� �� �� �� �� ��</strong></div></td>"
            Response.Write "    </tr>"
            Response.Write "  <tr class='tdbg'><td aling='center'><font color='#FF0000'>���ľۺ���δͨ�����!</font></td></tr>"
            Response.Write "</table>"
        End If
    Else
        Response.Write "<form method='post' action='User_Space.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
        Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� �� �� �� ��</strong></div></td>"
        Response.Write "    </tr>"
        Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>�ռ����ƣ�</strong><input name='BlogName' type='text' size='20' maxlength='20'> <font color='#FF0000'>* ��Ϊ�ۺϿռ����õ�����</font></td></tr>"
        Response.Write "  <tr class='tdbg'><td colspan='2'><table><tr><td>&nbsp;<strong>�ռ���ҳ<br>&nbsp;��ʾ��Ŀ��</strong></td><td>"
        Set rsBlogClass = Conn.Execute("select * from PE_Channel Where Disabled=" & PE_False & " and ModuleType>0 and ModuleType<4 and ChannelType=0 order by OrderID")
        Do While Not rsBlogClass.EOF
            If FoundInArr(arrClass_Input, rsBlogClass("ChannelDir") & "none", ",") = False Then
                Response.Write "<input type='checkbox' name='Showitem' value='" & rsBlogClass("ChannelID") & "' checked>����" & rsBlogClass("ChannelName") & "Ƶ���������Ʒ<br>"
            End If
            rsBlogClass.MoveNext
        Loop
        Response.Write "</td></tr></table></td></tr><tr class='tdbg'>"
        Response.Write "      <td height='22' colspan='2'><strong>ѡ����Ŀ</strong></td>"
        Response.Write "    </tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong>���ࣺ</strong><select name='BlogType'>" & GetKingOpti(0) & "</select></td>"
        Response.Write "    <td rowspan='10' align='center' valign='top' class='tdbg'>"
        Response.Write "        <table width='180' height='200' border='1'>"
        Response.Write "            <tr><td width='100%' align='center'><img id='img' src='" & InstallDir & "Space/default.gif' width='150' height='172'></td></tr>"
        Response.Write "        </table>"
        Response.Write "        <input name='url' type='text' size='25'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=UserBlogPic&size=" & UserSetting(27) & "' frameborder=0 scrolling=no width='285' height='25'></iframe>"
        Response.Write "     </td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>סַ��</strong><input name='Address' type='text' size='20' maxlength='20'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�绰��</strong><input name='Tel' type='text' size='20' maxlength='20'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>���棺</strong><input name='Fax' type='text' size='20' maxlength='20'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>��λ��</strong><input name='Company' type='text' size='20' maxlength='20'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>���ţ�</strong><input name='Department' type='text' size='20' maxlength='20'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�ѣѣ�</strong><input name='QQ' type='text'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�ʱࣺ</strong><input name='ZipCode' type='text' size='20' maxlength='20'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>��ҳ��</strong><input name='HomePage' type='text'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�ʼ���</strong><input name='Email' type='text' size='20' maxlength='20'></td></tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>�ۺϿռ���</strong>��<br>"
        Response.Write "      <textarea name='Intro' id='Intro' cols='72' rows='9' style='display:none'></textarea>"
        Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='300' ></iframe>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
        Response.Write "    <input name='addtype' type='hidden' id='addtype' value=1>"
        Response.Write "    <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
        Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Space.asp?Action=Manage';"" style='cursor:hand;'></td>"
        Response.Write "  </tr>"
        Response.Write "</table></form>"
    End If
Else
    Response.Write "<center>�����ڵ��û�����δ���žۺϿռ�!</center>"
End If
End Sub

Sub Modify()
    Dim BlogID
    Dim rsBlog, rsBlogClass, sqlBlog
    BlogID = Trim(Request("ID"))
    If BlogID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵľۺ�</li>"
        Exit Sub
    Else
        BlogID = PE_CLng(BlogID)
    End If
    sqlBlog = "Select * from PE_Space where ID=" & BlogID
    Set rsBlog = Server.CreateObject("Adodb.RecordSet")
    rsBlog.Open sqlBlog, Conn, 1, 3
    If rsBlog.BOF And rsBlog.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ˾ۺϣ�</li>"
    Else
        If rsBlog("type") > 1 Then
            Response.Write "<form method='post' action='User_Space.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
            Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
            Response.Write "    <tr class='title'> "
            Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� Ŀ �� Ϣ</strong></div></td>"
            Response.Write "    </tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Ŀ���ͣ�</strong><select name='addtype' onChange=""changemode()""><option value=2"
            If rsBlog("type") = 2 Then Response.Write " selected"
            Response.Write ">�ⲿRSS����</option><option value=3"
            If rsBlog("type") = 3 Then Response.Write " selected"
            Response.Write ">�ҵ���־</option><option value=4"
            If rsBlog("type") = 4 Then Response.Write " selected"
            Response.Write ">�ҵ�����</option><option value=5"
            If rsBlog("type") = 5 Then Response.Write " selected"
            Response.Write ">�ҵ�ͼ��</option><option value=6"
            If rsBlog("type") = 6 Then Response.Write " selected"
            Response.Write ">�ҵ����</option><option value=7"
            If rsBlog("type") = 7 Then Response.Write " selected"
            Response.Write ">�ҵ�����</option></select> <font color='#FF0000'>* ����Ŀʹ�õĿռ�ģ�� <a href='space_detal.html' target='_blank'>[�鿴��ϸ˵��]</a></font></td></tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��Ŀ���ƣ�</strong><input name='BlogName' type='text' size='45' value='" & rsBlog("Name") & "'> <font color='#FF0000'>*</font></td></tr>"
            If rsBlog("type") = 2 Then
                Response.Write "<tbody id='url' style='display:'>"
            Else
                Response.Write "<tbody id='url' style='display:none'>"
            End If
            Response.Write "  <tr class='tdbg' ><td colspan='2'>&nbsp;<strong>��Դ��ַ��</strong><input name='LinkUrl' type='text' size='45' maxlength='100' value='" & rsBlog("LinkUrl") & "'> <font color='#FF0000'>* ֧��RSS1.0 RSS2.0 ATOM��ʽ</font></td></tr></tbody>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>��ʾ������</strong><input name='ListNum' type='text' size='5' maxlength='3' value='" & rsBlog("listnum") & "'> <font color='#FF0000'>* ��ҳ���ҳ��ʾ����</font></td></tr>"
            Response.Write "  <tr class='tdbg'> "
            Response.Write "    <td colspan='2'>&nbsp;<strong>��ʾ����</strong>��<br>"
            Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>" & Server.HTMLEncode(rsBlog("Intro")) & "</textarea>"
            Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            Response.Write "    <tr>"
            Response.Write "      <td colspan='2' align='center' class='tdbg'>"
            Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
            Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsBlog("ID") & ">"
            Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Space.asp?Action=Manage&Passed=1';"" style='cursor:hand;'></td>"
            Response.Write "    </tr>"
            Response.Write "  </table>"
            Response.Write "</form>"
        Else
            Response.Write "<form method='post' action='User_Space.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
            Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
            Response.Write "    <tr class='title'> "
            Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� �� �� �� �� �� ��</strong></div></td>"
            Response.Write "    </tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>�ռ����ƣ�</strong><input name='BlogName' type='text' value='" & rsBlog("Name") & "'> <font color='#FF0000'>* ��Ϊ�ۺϿռ����õ�����</font></td></tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'><table><tr><td>&nbsp;<strong>�ռ���ҳ<br>&nbsp;��ʾ��Ŀ��</strong></td><td>"
            Set rsBlogClass = Conn.Execute("select * from PE_Channel Where Disabled=" & PE_False & " and ModuleType>0 and ModuleType<4 order by OrderID")
            Do While Not rsBlogClass.EOF
                If FoundInArr(arrClass_Input, rsBlogClass("ChannelDir") & "none", ",") = False Then
                    Response.Write "<input type='checkbox' name='Showitem' value='" & rsBlogClass("ChannelID") & "'"
                    If FoundInArr(rsBlog("LinkUrl"), rsBlogClass("ChannelID"), ",") Then Response.Write " checked"
                    Response.Write ">����" & rsBlogClass("ChannelName") & "Ƶ���������Ʒ<br>"
                End If
                rsBlogClass.MoveNext
            Loop
            Response.Write "</td></tr></table></td></tr><tr class='tdbg'>"

            Response.Write "    <tr class='title'> "
            Response.Write "      <td height='22' colspan='2'><strong>ѡ����Ŀ</strong></td>"
            Response.Write "    </tr>"
            Response.Write "  <tr class='tdbg'> "
            Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong>���ࣺ</strong><select name='BlogType'>" & GetKingOpti(rsBlog("ClassID")) & "</select></td>"
            Response.Write "    <td rowspan='10' align='center' valign='top' class='tdbg'>"
            Response.Write "        <table width='180' height='200' border='1'>"
            Response.Write "            <tr><td width='100%' align='center'>"
            If Trim(rsBlog("Photo") & "") = "" Then
                Response.Write "<img id='img' src='" & InstallDir & "Space/default.gif' width='150' height='172'>"
            Else
                Response.Write "<img id='img' src='" & rsBlog("Photo") & "' width='150' height='172'>"
            End If
            Response.Write "        </td></tr></table>"
            Response.Write "        <input name='url' type='text' size='25' value='" & rsBlog("Photo") & "'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=UserBlogPic&size=" & UserSetting(27) & "' frameborder=0 scrolling=no width='285' height='25'></iframe>"
            Response.Write "     </td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>��ַ��</strong><input name='Address' type='text'  value='" & rsBlog("Address") & "'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�绰��</strong><input name='Tel' type='text' value='" & rsBlog("Tel") & "'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>���棺</strong><input name='Fax' type='text' value='" & rsBlog("Fax") & "'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>��λ��</strong><input name='Company' type='text' value='" & rsBlog("Company") & "'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>���ţ�</strong><input name='Department' type='text' value='" & rsBlog("Department") & "'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�ʱࣺ</strong><input name='ZipCode' type='text' value='" & rsBlog("ZipCode") & "'></td>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�ѣѣ�</strong><input name='QQ' type='text' value='" & rsBlog("QQ") & "'></td>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>��ҳ��</strong><input name='HomePage' type='text' value='" & rsBlog("HomePage") & "'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>�ʼ���</strong><input name='Email' type='text' value='" & rsBlog("Email") & "'></td></tr>"
            Response.Write "  <tr>"
            Response.Write "  <tr class='tdbg'> "
            Response.Write "    <td colspan='2'>&nbsp;<strong>�ۺϿռ���</strong>��<br>"
            Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>" & Server.HTMLEncode(rsBlog("Intro")) & "</textarea>"
            Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            Response.Write "    <tr>"
            Response.Write "      <td colspan='2' align='center' class='tdbg'>"
            Response.Write "    <input name='addtype' type='hidden' id='addtype' value=1>"
            Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
            Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsBlog("ID") & ">"
            Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Space.asp?Action=Manage&Passed=1';"" style='cursor:hand;'></td>"
            Response.Write "    </tr>"
            Response.Write "  </table>"
            Response.Write "</form>"
        End If
    End If
    rsBlog.Close
    Set rsBlog = Nothing
End Sub

Sub SaveBlog()
    Dim BlogID, BlogType, BlogName, Address, Tel, Fax, Company, Department, ZipCode, Homepage, Email, QQ, Intro, Photo, LinkUrl
    Dim rsBlog, sqlBlog, isFirst, addtype, listnum
    isFirst = False

    BlogName = Trim(Request.Form("BlogName"))
    If BlogName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���Ʋ���Ϊ�գ�</li>"
    Else
        BlogName = ReplaceBadChar(BlogName)
    End If

    If Action = "SaveModify" Then
        BlogID = Trim(Request.Form("ID"))
        If BlogID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����������</li>"
        Else
            BlogID = PE_CLng(BlogID)
        End If
    End If
    
    Dim cusers, UserPassword, LastPassword
    UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
    LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))
    If Action = "SaveAdd" Then
        Set cusers = Conn.Execute("select UserID,UserName,UserPassword,LastPassword from PE_User Where UserID=" & UserID)
    Else
        Set cusers = Conn.Execute("select A.ID,C.UserID,C.UserName,C.UserPassword,C.LastPassword from PE_Space A inner join PE_User C on A.UserID=C.UserID Where A.ID=" & BlogID)
    End If
    If cusers.BOF And cusers.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����֤ʧ�ܣ�</li>"
    Else
        If UserName <> cusers("UserName") Or UserPassword <> cusers("UserPassword") Or LastPassword <> cusers("LastPassword") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����֤ʧ�ܣ�</li>"
        End If
    End If
    Set cusers = Nothing

    If FoundErr = True Then Exit Sub

    addtype = PE_CLng(Trim(Request.Form("addtype")))

    If addtype = 0 Or addtype = 2 Then
        addtype = 2
        LinkUrl = Trim(Request.Form("LinkUrl"))
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��Դ����Ϊ�գ�</li>"
        Else
            Dim XmlRss, RssDOM, oItem, tetitle, teurl
            On Error Resume Next
            Set XmlRss = Server.CreateObject("MSXML2.ServerXMLHTTP")
            XmlRss.SetTimeouts 5000, 5000, 120000, 60000
            XmlRss.Open "GET", LinkUrl, False
            XmlRss.Send
            If Err.Number <> 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��Դ��ַ�����ڻ��޷����ӣ�</li>"
                Err.Clear
            Else
                If XmlRss.Readystate <> 4 Or Trim(XmlRss.responseText & "") = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��Դ��ַ�����ڻ��޷����ӣ�</li>"
                Else
                    Set RssDOM = Server.CreateObject("microsoft.XMLDOM")
                    RssDOM.async = False
                    RssDOM.Load (XmlRss.responseXML)
                    If RssDOM.Readystate <> 4 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>��Դ��ַ������Ч��XML����Դ</li>"
                    End If
                    Set RssDOM = Nothing
                End If
            End If
            Set XmlRss = Nothing
        End If
    ElseIf addtype = 1 Then
        LinkUrl = Trim(Request.Form("Showitem"))
        Photo = PE_HTMLEncode(Trim(Request.Form("url")))
        Address = PE_HTMLEncode(Trim(Request.Form("Address")))
        Tel = PE_HTMLEncode(Trim(Request.Form("Tel")))
        Fax = PE_HTMLEncode(Trim(Request.Form("Fax")))
        Company = PE_HTMLEncode(Trim(Request.Form("Company")))
        Department = PE_HTMLEncode(Trim(Request.Form("Department")))
        ZipCode = PE_HTMLEncode(Trim(Request.Form("ZipCode")))
        Homepage = PE_HTMLEncode(Trim(Request.Form("HomePage")))
        Email = PE_HTMLEncode(Trim(Request.Form("Email")))
        QQ = PE_HTMLEncode(Trim(Request.Form("QQ")))
    End If

    BlogType = PE_CLng(Trim(Request.Form("BlogType")))
    listnum = PE_CLng(Trim(Request.Form("ListNum")))
    If listnum = 0 Then listnum = 10
    Intro = ReplaceBadUrl(Trim(Request.Form("Intro")))

    If FoundErr = True Then Exit Sub

    If Action = "SaveAdd" Then
        BlogID = PE_CLng(Conn.Execute("select max(ID) from PE_Space")(0)) + 1

        Set rsBlog = Conn.Execute("Select Top 1 UserID,Passed from PE_Space where UserID=" & UserID & " and Type=1")
        If rsBlog.BOF And rsBlog.EOF Then
            isFirst = True
            Conn.Execute ("update PE_User set Blog=" & PE_True & " where UserID=" & UserID)
        Else
            If rsBlog("Passed") = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���ľۺ���δͨ������,���������Ŀ��</li>"
                Set rsBlog = Nothing
                Call CloseConn
                Exit Sub
            End If
        End If

        Set rsBlog = Server.CreateObject("Adodb.RecordSet")
        sqlBlog = "Select * from PE_Space"
        rsBlog.Open sqlBlog, Conn, 1, 3
        rsBlog.AddNew
        rsBlog("ID") = BlogID
        rsBlog("UserID") = UserID
        rsBlog("ClassID") = BlogType
        rsBlog("Name") = BlogName
        rsBlog("BirthDay") = Now()
        If addtype = 1 Then
            rsBlog("Address") = Address
            rsBlog("Tel") = Tel
            rsBlog("Fax") = Fax
            rsBlog("Company") = Company
            rsBlog("Department") = Department
            rsBlog("ZipCode") = ZipCode
            rsBlog("HomePage") = Homepage
            rsBlog("Email") = Email
            rsBlog("QQ") = PE_CLng(QQ)
        End If
        rsBlog("Intro") = FilterJS(Intro)
        If Photo <> "" Then rsBlog("Photo") = Photo

        If isFirst = True Then
            rsBlog("Type") = 1
            rsBlog("OrderID") = 1
            If PE_CLng(UserSetting(26)) = 1 Then
                rsBlog("Passed") = True
            Else
                rsBlog("Passed") = False
            End If
        Else
            rsBlog("Type") = addtype
            rsBlog("OrderID") = 2
            rsBlog("Passed") = True
        End If
        rsBlog("LastUseTime") = Now()
        If Trim(LinkUrl & "") = "" Then
           rsBlog("LinkUrl") = Null
        Else
           rsBlog("LinkUrl") = LinkUrl
        End If
        rsBlog("listnum") = listnum
        rsBlog.Update
        If addtype = 1 And isFirst = True Then Call CreateBlogDir(UserID, UserName)
    Else
        Set rsBlog = Server.CreateObject("Adodb.RecordSet")
        sqlBlog = "Select * from PE_Space where ID=" & BlogID
        rsBlog.Open sqlBlog, Conn, 1, 3
        If rsBlog.BOF And rsBlog.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>δ�ҵ�����ۺϣ�</li>"
            rsBlog.Close
            Set rsBlog = Nothing
            Exit Sub
        End If
            rsBlog("Name") = BlogName
            rsBlog("ClassID") = BlogType
            If addtype = 1 Then
                rsBlog("Address") = Address
                rsBlog("Tel") = Tel
                rsBlog("Fax") = Fax
                rsBlog("Company") = Company
                rsBlog("Department") = Department
                rsBlog("ZipCode") = ZipCode
                rsBlog("HomePage") = Homepage
                rsBlog("Email") = Email
                rsBlog("QQ") = PE_CLng(QQ)
            End If
            rsBlog("Intro") = Intro
            If Photo <> "" Then rsBlog("Photo") = Photo
            rsBlog("Type") = addtype
            If Trim(LinkUrl & "") = "" Then
               rsBlog("LinkUrl") = Null
            Else
               rsBlog("LinkUrl") = LinkUrl
            End If
            rsBlog("listnum") = listnum
            rsBlog.Update
    End If
    rsBlog.Close
    Set rsBlog = Nothing
    Call CloseConn
    Response.Redirect "User_Space.asp?Action=Manage"
End Sub

Sub Del()
    Dim BlogID, cusers, UserPassword, LastPassword
    BlogID = PE_CLng(Trim(Request("ID")))
    If BlogID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ���ۺϣ�</li>"
        Exit Sub
    End If

    UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
    LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))
    Set cusers = Conn.Execute("select A.ID,C.UserID,C.UserName,C.UserPassword,C.LastPassword from PE_Space A inner join PE_User C on A.UserID=C.UserID Where A.ID=" & BlogID)
    If cusers.BOF And cusers.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����֤ʧ�ܣ�</li>"
    Else
        If UserName <> cusers("UserName") Or UserPassword <> cusers("UserPassword") Or LastPassword <> cusers("LastPassword") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����֤ʧ�ܣ�</li>"
        End If
    End If
    Set cusers = Nothing
    If FoundErr = True Then Exit Sub

    Conn.Execute ("delete from PE_Space where ID=" & CLng(BlogID) & "")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub SetStat()
    Dim cusers, UserPassword, LastPassword, BlogID, OrderID, tmporderid
    tmporderid = Split(Action, "|")
    If UBound(tmporderid) = 1 Then
        BlogID = PE_CLng(tmporderid(1))
        OrderID = PE_CLng(Trim(Request.Form("OrderID" & BlogID)))
        UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
        LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))
        Set cusers = Conn.Execute("select A.ID,C.UserID,C.UserName,C.UserPassword,C.LastPassword from PE_Space A inner join PE_User C on A.UserID=C.UserID Where A.ID=" & BlogID)
        If cusers.BOF And cusers.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����֤ʧ�ܣ�</li>"
        Else
            If UserName <> cusers("UserName") Or UserPassword <> cusers("UserPassword") Or LastPassword <> cusers("LastPassword") Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�����֤ʧ�ܣ�</li>"
            End If
        End If
        Set cusers = Nothing
        If FoundErr = True Then Exit Sub
        If OrderID > 1 And BlogID > 0 Then Conn.Execute ("update PE_Space set OrderID=" & OrderID & " where ID=" & BlogID & "")
    End If
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub CreateBlogDir(Uid, UName)
    If PE_CLng(Uid) = 0 Or Trim(UName & "") = "" Then Exit Sub
    On Error Resume Next
    Dim fsfl, fl, strDir
    
    'ǿ��ʹ���û�ID��β����ֹ�����Ƿ�Ŀ¼
    Dim spacename
    spacename = Replace(LCase(UName & Uid), ".", "")

    strDir = InstallDir & "Space/" & spacename & "/"
    If fso.FolderExists(Server.MapPath(strDir)) = False Then fso.CreateFolder Server.MapPath(strDir)

    Set fsfl = fso.GetFolder(Server.MapPath(InstallDir & "Space/Default/"))
    For Each fl In fsfl.Files
        fl.Copy Server.MapPath(strDir & fl.name), True
    Next

    Set fsfl = fso.CreateTextFile(Server.MapPath(strDir & "config.xml"), True)
    fsfl.WriteLine ("<?" & "xml version=""1.0"" encoding=""gb2312""" & "?>")
    fsfl.WriteLine ("<" & "body" & ">")
    fsfl.WriteLine ("<" & "baseconfig" & ">")
    fsfl.WriteLine ("<" & "userid" & ">" & Uid & "</" & "userid" & ">")
    fsfl.WriteLine ("</" & "baseconfig" & ">")
    fsfl.WriteLine ("</" & "body" & ">")

    '���þۺ�Ϊδ���״̬
    If PE_CLng(UserSetting(26)) = 0 Then
        Set fsfl = fso.CreateTextFile(Server.MapPath(strDir & "index.asp"), True)
        fsfl.WriteLine ("�����...")
    End If
    fsfl.Close
    Set fsfl = Nothing
End Sub

Sub Template()
If PE_CLng(UserSetting(28)) = 1 Then
    On Error Resume Next
    Dim fsfl, fc, fl, UDir
    Dim spacename
    spacename = Replace(LCase(UserName & UserID), ".", "")

    UDir = InstallDir & "Space/" & spacename & "/"
    If fso.FolderExists(Server.MapPath(UDir)) = False Then
        Response.Write "<br><center>�û��ռ䲻����<br><br><a href='User_Space.asp?Action=Template'>�� ���� ��</a></center>"
    Else
        Response.Write "<br>"
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td align='center'>��ѡ������ʹ�õĿռ�Ƥ��</td></tr></table>"
        Response.Write "<table class='border' border='0' cellspacing='15' width='100%' cellpadding='15'><tr align='center'>"
        Set fsfl = fso.GetFolder(Server.MapPath(InstallDir & "Space/Template/"))
        Set fc = fsfl.SubFolders
        i = 1
        For Each fl In fc
            Response.Write "<td><a href='#' onclick=""changetemplate('" & fl.name & "');""><img src='" & InstallDir & "Space/Template/" & fl.name & ".gif' border='0' alt='" & fl.name & "'><br>" & fl.name & "</a></td>"
            If i Mod 3 = 0 Then
                Response.Write "</tr><tr align='center'>"
            End If
            i = i + 1
        Next
        Set fsfl = Nothing
        Response.Write "</table>"
    End If
Else
    Response.Write "<br><center>����Ȩ�����ռ�Ƥ��<br><br><a href='User_Space.asp?Action=Manage'>�� ���� ��</a></center>"
End If
End Sub

Sub CTemplate()
If PE_CLng(UserSetting(28)) = 1 Then
    Dim fname
    fname = Trim(Request("fname"))
    fname = Replace(Replace(fname, ".", ""), "/", "")
    Dim fsfl, fl, UDir, spacename
    spacename = Replace(LCase(UserName & UserID), ".", "")
    UDir = InstallDir & "Space/" & spacename & "/"
    If fso.FolderExists(Server.MapPath(UDir)) = False Then
        Response.Write "<br><center>�û��ռ䲻����<br><br><a href='User_Space.asp?Action=Template'>�� ���� ��</a></center>"
    Else
        dim fflag, fc
        fflag = 0
        Set fsfl = fso.GetFolder(Server.MapPath(InstallDir & "Space/Template/"))
        Set fc = fsfl.SubFolders
        For Each fl In fc
            If fl.name = fname Then
                fflag = 1
            End If
        Next
        If fflag = 1 Then
            Set fsfl = fso.GetFolder(Server.MapPath(InstallDir & "Space/Template/" & fname))
            For Each fl In fsfl.Files
                fl.Copy Server.MapPath(UDir & fl.name), True
            Next
            If fso.FolderExists(Server.MapPath(InstallDir & "Space/Template/" & fname & "/skin")) Then
                If fso.FolderExists(Server.MapPath(UDir & "skin")) Then
                    fso.DeleteFolder (Server.MapPath(UDir & "skin"))
                End If
                fso.CopyFolder Server.MapPath(InstallDir & "Space/Template/" & fname & "/skin"), Server.MapPath(UDir & "skin")
            End If
            Response.Write "<br><center>���Ŀռ��Ѿ��ɹ���Ӧ������Ƥ����" & fname & "��!<br><br><a href='../Space/" & spacename & "' target='_blank'>�� �鿴Ч�� ��</a><br><a href='User_Space.asp?Action=Template'>�� ���� ��</a></center>"
        Else
            Response.Write "<br><center>��ѡ���Ƥ��������<br><br><a href='User_Space.asp?Action=Manage'>�� ���� ��</a></center>"       
        End If
        Set fsfl = Nothing

    End If
Else
    Response.Write "<br><center>����Ȩ�����ռ�Ƥ��<br><br><a href='User_Space.asp?Action=Manage'>�� ���� ��</a></center>"
End If
End Sub

Function GetKingOpti(iselected)
    Dim strtmp, rskind
    Set rskind = Conn.Execute("select KindID,KindName from PE_SpaceKind order by OrderID")
    Do While Not rskind.EOF
        strtmp = strtmp & "<option value=" & rskind("KindID")
        If iselected = rskind("KindID") Then
            strtmp = strtmp & " selected"
        End If
        strtmp = strtmp & ">" & rskind("KindName") & "</option>"
        rskind.MoveNext
    Loop
    Set rskind = Nothing
    strtmp = strtmp & "<option value=0"
    If iselected = 0 Then
        strtmp = strtmp & " selected"
    End If
    strtmp = strtmp & ">�������κη���</option>"
    GetKingOpti = strtmp
End Function

Function GetKingName(iselected)
    Dim strtmp, rskind, KindS

    If oldKInd = "" Then oldKInd = "0|||�޷���"

    KindS = Split(oldKInd, "|||")
    If KindS(0) <> iselected Then
        Set rskind = Conn.Execute("select top 1 KindID,KindName from PE_SpaceKind Where KindID=" & iselected)
        If Not (rskind.BOF And rskind.EOF) Then
            strtmp = rskind("KindName")
        Else
            strtmp = "�޷���"
        End If
        oldKInd = iselected & "|||" & strtmp
        Set rskind = Nothing
    Else
        strtmp = KindS(1)
    End If
    GetKingName = strtmp
End Function
%>
