<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_JS.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Sub Add()
    Response.Write "<form action='Admin_PhotoJS.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>����µ�JS�ļ�����ͨ�б�ʽ��</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif("", "", 0, "")

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowType' id='ShowType'>"
    Response.Write "          <option value='1' selected>��ͨ�б�</option>"
    Response.Write "          <option value='2'>���ʽ</option>"
    Response.Write "          <option value='3'>�������ʽ</option>"
    Response.Write "          <option value='4'>DIV���</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ߣ�</strong></td>"
    Response.Write "      <td height='25'><input name='Author' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>�����Ϊ�գ���ֻ��ʾָ�����ߵ�" & ChannelShortName & "�����ڸ���" & ChannelShortName & "����</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "��Ŀ��</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='10' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>������Ŀ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>������Ŀ</option>" & GetClass_Option(0) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' value='1' checked>��������Ŀ&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>����ָ��Ϊ�ⲿ��Ŀ"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>����ר�⣺</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>�������κ�ר��</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "���ԣ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' value='1'> ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' value='1'> �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>�������ѡ������ʾ����" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ڷ�Χ��</strong></td>"
    Response.Write "      <td height='25'>ֻ��ʾ��� <input name='DateNum' type='text' id='DateNum' value='30' size='5' maxlength='3'> ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ����������" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���򷽷���</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option value='1' selected>" & ChannelShortName & "ID������</option>"
    Response.Write "          <option value='2'>" & ChannelShortName & "ID������</option>"
    Response.Write "          <option value='3'>����ʱ�䣨����</option>"
    Response.Write "          <option value='4'>����ʱ�䣨����</option>"
    Response.Write "          <option value='5'>�������������</option>"
    Response.Write "          <option value='6'>�������������</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>����ͼƬ��ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowPropertyType' id='ShowPropertyType'>"
    Response.Write "          <option value='0' >����ʾ</option>"
    Response.Write "          <option value='2'>����</option>"
    Response.Write "          <option value='1' selected>СͼƬ����ʽ1��</option>"
    Response.Write "          <option value='3' selected>СͼƬ����ʽ2��</option>"
    Response.Write "          <option value='4' selected>СͼƬ����ʽ3��</option>"
    Response.Write "          <option value='5' selected>СͼƬ����ʽ4��</option>"
    Response.Write "          <option value='6' selected>СͼƬ����ʽ5��</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "�����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='30' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>���Ϊ0������ʾ�������ơ���ĸ��һ���ַ��������������ַ���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='0' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������0������" & ChannelShortName & "�����·���ʾָ��������" & ChannelShortName & "���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='50' align='right'><strong>��ʾ���ݣ�</strong></td>"
    Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowClassName' type='checkbox' id='ShowClassName' value='1'>������Ŀ</td>"
    Response.Write "          <td><input name='ShowAuthor' type='checkbox' id='ShowAuthor' value='1'>����</td>"
    Response.Write "          <td>����ʱ��"
    Response.Write "            <select name='ShowDateType' id='ShowDateType'>"
    Response.Write "              <option value='0'>����ʾ</option>"
    Response.Write "              <option value='1'>������</option>"
    Response.Write "              <option value='2'>����</option>"
    Response.Write "              <option value='3'>��-��</option>"
    Response.Write "            </select>"
    Response.Write "          </td>"
    Response.Write "          <td><input name='ShowHits' type='checkbox' id='ShowHits' value='1' checked>�������</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowHotSign' type='checkbox' id='ShowHotSign' value='1'>����" & ChannelShortName & "��־</td>"
    Response.Write "          <td><input name='ShowNewSign' type='checkbox' id='ShowNewSign' value='1'>����" & ChannelShortName & "��־</td>"
    Response.Write "          <td><input name='ShowTips' type='checkbox' id='ShowTips' value='1'>��ʾ��ʾ��Ϣ</td>"
    Response.Write "          <td> </td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "�򿪷�ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OpenType' id='OpenType'>"
    Response.Write "          <option value='0'>��ԭ���ڴ�</option>"
    Response.Write "          <option value='1' selected>���´��ڴ�</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ӵ�ַѡ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option value='0' selected>ʹ�����·��</option>"
    Response.Write "          <option value='1'>ʹ�ð���������ַ�ľ���·��</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>ÿ�б���������</strong></td>"
    Response.Write "      <td height='25'><input name='Cols' type='text' value='1' size='5' maxlength='3'> <font color='#FF0000'>ÿ����ʾ���������</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>CSS���������</strong></td>"
    Response.Write "      <td height='25'><input name='CssNameA' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>�б����������ӵ��õ�CSS����</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>�����ʽ1��</strong></td>"
    Response.Write "      <td height='25'><input name='CssName1' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>�б��������е�CSSЧ��������</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>�����ʽ2��</strong></td>"
    Response.Write "      <td height='25'><input name='CssName2' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>�б���ż���е�CSSЧ��������</font></td>"
    Response.Write "    </tr>"

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim ID, sqlJs, rsJs, JsConfig
    ID = PE_CLng(Trim(Request("ID")))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
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
    JsConfig = Split(rsJs("Config") & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0", "|")

    Response.Write "<form action='Admin_PhotoJS.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�޸Ĳ�������ͨ�б�ʽ��</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif(rsJs("JsName"), rsJs("JsReadme"), rsJs("ContentType"), rsJs("JsFileName"))

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowType' id='ShowType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 1) & ">��ͨ�б�</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 2) & ">���ʽ</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 3) & ">�������ʽ</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 4) & ">DIV���</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ߣ�</strong></td>"
    Response.Write "      <td height='25'><input name='Author' type='text' value='" & ZeroToEmpty(JsConfig(21)) & "' size='10' maxlength='20'> <font color='#FF0000'>�����Ϊ�գ���ֻ��ʾָ�����ߵ�" & ChannelShortName & "�����ڸ���" & ChannelShortName & "����</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "��Ŀ��</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='" & JsConfig(1) & "' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>������Ŀ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>������Ŀ</option>" & GetClass_Option(PE_CLng(JsConfig(2))) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' " & RadioValue(PE_CLng(JsConfig(3)), 1) & ">��������Ŀ&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>����ָ��Ϊ�ⲿ��Ŀ"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>����ר�⣺</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>�������κ�ר��</option>" & GetSpecial_Option(PE_CLng(JsConfig(20))) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "���ԣ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' " & RadioValue(PE_CLng(JsConfig(4)), 1) & ">����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' " & RadioValue(PE_CLng(JsConfig(5)), 1) & ">�Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>�������ѡ������ʾ����" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ڷ�Χ��</strong></td>"
    Response.Write "      <td height='25'>ֻ��ʾ��� <input name='DateNum' type='text' id='DateNum' value='" & JsConfig(6) & "' size='5' maxlength='3'> ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ����������" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���򷽷���</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 1) & ">" & ChannelShortName & "ID������</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 2) & ">" & ChannelShortName & "ID������</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 3) & ">����ʱ�䣨����</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 4) & ">����ʱ�䣨����</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 5) & ">�������������</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 6) & ">�������������</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>����ͼƬ��ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowPropertyType' id='ShowPropertyType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 0) & ">����ʾ</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 2) & ">����</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 1) & ">СͼƬ����ʽ1��</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 3) & ">СͼƬ����ʽ2��</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 4) & ">СͼƬ����ʽ3��</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 5) & ">СͼƬ����ʽ4��</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 6) & ">СͼƬ����ʽ5��</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "�����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='" & JsConfig(8) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>���Ϊ0������ʾ�������ơ���ĸ��һ���ַ��������������ַ���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='" & JsConfig(9) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������0������" & ChannelShortName & "�����·���ʾָ��������" & ChannelShortName & "���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='50' align='right'><strong>��ʾ���ݣ�</strong></td>"
    Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowClassName' type='checkbox' id='ShowClassName' " & RadioValue(PE_CLng(JsConfig(10)), 1) & ">������Ŀ</td>"
    Response.Write "          <td><input name='ShowAuthor' type='checkbox' id='ShowAuthor' " & RadioValue(PE_CLng(JsConfig(11)), 1) & ">����</td>"
    Response.Write "          <td>����ʱ��"
    Response.Write "            <select name='ShowDateType' id='ShowDateType'>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 0) & ">����ʾ</option>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 1) & ">������</option>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 2) & ">����</option>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 3) & ">��-��</option>"
    Response.Write "            </select>"
    Response.Write "          </td>"
    Response.Write "          <td><input name='ShowHits' type='checkbox' id='ShowHits' " & RadioValue(PE_CLng(JsConfig(13)), 1) & ">�������</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowHotSign' type='checkbox' id='ShowHotSign' " & RadioValue(PE_CLng(JsConfig(14)), 1) & ">����" & ChannelShortName & "��־</td>"
    Response.Write "          <td><input name='ShowNewSign' type='checkbox' id='ShowNewSign' " & RadioValue(PE_CLng(JsConfig(15)), 1) & ">����" & ChannelShortName & "��־</td>"
    Response.Write "          <td><input name='ShowTips' type='checkbox' id='ShowTips' " & RadioValue(PE_CLng(JsConfig(16)), 1) & ">��ʾ��ʾ��Ϣ</td>"
    Response.Write "          <td> </td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "�򿪷�ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OpenType' id='OpenType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(17)), 0) & ">��ԭ���ڴ�</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(17)), 1) & ">���´��ڴ�</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ӵ�ַѡ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(18)), 0) & ">ʹ�����·��</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(18)), 1) & ">ʹ�ð���������ַ�ľ���·��</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>ÿ�б���������</strong></td>"
    Response.Write "      <td height='25'><input name='Cols' type='text' value='" & PE_CLng(JsConfig(22)) & "' size='5' maxlength='3'> <font color='#FF0000'>ÿ����ʾ���������</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>CSS���������</strong></td>"
    Response.Write "      <td height='25'><input name='CssNameA' type='text' value='" & ZeroToEmpty(JsConfig(23)) & "' size='10' maxlength='20'> <font color='#FF0000'>�б����������ӵ��õ�CSS����</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>�����ʽ1��</strong></td>"
    Response.Write "      <td height='25'><input name='CssName1' type='text' value='" & ZeroToEmpty(JsConfig(24)) & "' size='10' maxlength='20'> <font color='#FF0000'>�б��������е�CSSЧ��������</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>�����ʽ2��</strong></td>"
    Response.Write "      <td height='25'><input name='CssName2' type='text' value='" & ZeroToEmpty(JsConfig(25)) & "' size='10' maxlength='20'> <font color='#FF0000'>�б���ż���е�CSSЧ��������</font></td>"
    Response.Write "    </tr>"

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='ID' type='hidden' id='ID' value='" & ID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='�����޸Ľ��'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsJs.Close
    Set rsJs = Nothing
End Sub

Sub AddPic()
    Response.Write "<form action='Admin_PhotoJS.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>����µ�JS�ļ���ͼƬ�б�ʽ��</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif("", "", 0, "")

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>������Ŀ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>������Ŀ</option>" & GetClass_Option(0) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' value='1' checked>��������Ŀ&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>����ָ��Ϊ�ⲿ��Ŀ"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>����ר�⣺</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>�������κ�ר��</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "��Ŀ��</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='4' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "���ԣ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' value='1'> ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' value='1'> �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>�������ѡ������ʾ����" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ڷ�Χ��</strong></td>"
    Response.Write "      <td height='25'>ֻ��ʾ��� <input name='DateNum' type='text' id='DateNum' value='30' size='5' maxlength='3'> ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ����������" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���򷽷���</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option value='1' selected>" & ChannelShortName & "ID������</option>"
    Response.Write "          <option value='2'>" & ChannelShortName & "ID������</option>"
    Response.Write "          <option value='3'>����ʱ�䣨����</option>"
    Response.Write "          <option value='4'>����ʱ�䣨����</option>"
    Response.Write "          <option value='5'>�������������</option>"
    Response.Write "          <option value='6'>�������������</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='ShowType' type='radio' value='1' checked> ͼƬ+����+���ݼ�飺��������<br>"
    Response.Write "        <input name='ShowType' type='radio' value='2'> ��ͼƬ+���ƣ��������У�+���ݼ�飺��������<br>"
    Response.Write "        <input name='ShowType' type='radio' value='3'> ͼƬ+������+���ݼ�飺�������У�����������"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><b>��ҳͼƬ���ã�</b></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        ��ȣ� <input name='ImgWidth' type='text' id='ImgWidth' value='130' size='5' maxlength='3'> ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        �߶ȣ� <input name='ImgHeight' type='text' id='ImgHeight' value='90' size='5' maxlength='3'> ����"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "�����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='20' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>��Ϊ0������ʾ���ƣ���Ϊ-1������ʾ�������ơ���ĸ��һ���ַ��������������ַ���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='0' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������0������ʾָ��������" & ChannelShortName & "���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>��ʾ���ݣ�</strong></td>"
    Response.Write "      <td height='25'><input name='ShowTips' type='checkbox' id='ShowTips' value='1'> ��ʾ���ߡ�����ʱ�䡢���������ʾ��Ϣ</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>ÿ����ʾ" & ChannelShortName & "����</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='Cols' id='Cols'>"
    Response.Write "          <option value='1'>1</option>"
    Response.Write "          <option value='2'>2</option>"
    Response.Write "          <option value='3'>3</option>"
    Response.Write "          <option value='4' selected>4</option>"
    Response.Write "          <option value='5'>5</option>"
    Response.Write "          <option value='6'>6</option>"
    Response.Write "          <option value='7'>7</option>"
    Response.Write "          <option value='8'>8</option>"
    Response.Write "          <option value='9'>9</option>"
    Response.Write "          <option value='10'>10</option>"
    Response.Write "          <option value='11'>11</option>"
    Response.Write "          <option value='12'>12</option>"
    Response.Write "        </select>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;����ָ�������ͻỻ��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ӵ�ַѡ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option value='0' selected>ʹ�����·��</option>"
    Response.Write "          <option value='1'>ʹ�ð���������ַ�ľ���·��</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAddPic'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyPic()
    Dim ID, sqlJs, rsJs, JsConfig
    ID = PE_CLng(Trim(Request("ID")))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
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
    JsConfig = Split(rsJs("Config") & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0", "|")

    Response.Write "<form action='Admin_PhotoJS.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�޸Ĳ�����ͼƬ�б�ʽ��</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif(rsJs("JsName"), rsJs("JsReadme"), rsJs("ContentType"), rsJs("JsFileName"))

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>������Ŀ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>������Ŀ</option>" & GetClass_Option(PE_CLng(JsConfig(0))) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' " & RadioValue(PE_CLng(JsConfig(1)), 1) & ">��������Ŀ&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>����ָ��Ϊ�ⲿ��Ŀ"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>����ר�⣺</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>�������κ�ר��</option>" & GetSpecial_Option(PE_CLng(JsConfig(15))) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "��Ŀ��</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='" & JsConfig(2) & "' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "���ԣ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' " & RadioValue(PE_CLng(JsConfig(3)), 1) & ">����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' " & RadioValue(PE_CLng(JsConfig(4)), 1) & ">�Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>�������ѡ������ʾ����" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ڷ�Χ��</strong></td>"
    Response.Write "      <td height='25'>ֻ��ʾ��� <input name='DateNum' type='text' id='DateNum' value='" & JsConfig(5) & "' size='5' maxlength='3'> ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ����������" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���򷽷���</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 1) & ">" & ChannelShortName & "ID������</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 2) & ">" & ChannelShortName & "ID������</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 3) & ">����ʱ�䣨����</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 4) & ">����ʱ�䣨����</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 5) & ">�������������</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 6) & ">�������������</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='ShowType' type='radio' " & RadioValue(PE_CLng(JsConfig(7)), 1) & ">ͼƬ+����+���ݼ�飺��������<br>"
    Response.Write "        <input name='ShowType' type='radio' " & RadioValue(PE_CLng(JsConfig(7)), 2) & ">��ͼƬ+���ƣ��������У�+���ݼ�飺��������<br>"
    Response.Write "        <input name='ShowType' type='radio' " & RadioValue(PE_CLng(JsConfig(7)), 3) & ">ͼƬ+������+���ݼ�飺�������У�����������"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><b>��ҳͼƬ���ã�</b></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        ��ȣ� <input name='ImgWidth' type='text' id='ImgWidth' value='" & JsConfig(8) & "' size='5' maxlength='3'> ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        �߶ȣ� <input name='ImgHeight' type='text' id='ImgHeight' value='" & JsConfig(9) & "' size='5' maxlength='3'> ����"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "�����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='" & JsConfig(10) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>��Ϊ0������ʾ���ƣ���Ϊ-1������ʾ�������ơ���ĸ��һ���ַ��������������ַ���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='" & JsConfig(11) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������0������ʾָ��������" & ChannelShortName & "���</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>��ʾ���ݣ�</strong></td>"
    Response.Write "      <td height='25'><input name='ShowTips' type='checkbox' id='ShowTips' " & RadioValue(PE_CLng(JsConfig(12)), 1) & ">��ʾ���ߡ�����ʱ�䡢���������ʾ��Ϣ</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>ÿ����ʾ" & ChannelShortName & "����</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='Cols' id='Cols'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 1) & ">1</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 2) & ">2</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 3) & ">3</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 4) & ">4</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 5) & ">5</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 6) & ">6</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 7) & ">7</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 8) & ">8</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 9) & ">9</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 10) & ">10</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 11) & ">11</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(13)), 12) & ">12</option>"
    Response.Write "        </select>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;����ָ�������ͻỻ��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>���ӵ�ַѡ�</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(14)), 0) & ">ʹ�����·��</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(14)), 1) & ">ʹ�ð���������ַ�ľ���·��</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='ID' type='hidden' id='ID' value='" & ID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModifyPic'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='�����޸Ľ��'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsJs.Close
    Set rsJs = Nothing
End Sub

Sub SaveJS_List()
    Dim ID, JsName, JsReadme, JsFileName, Config
    Dim ShowType, PhotoNum, ClassID, SpecialID, IncludeChild, IsHot, IsElite, DateNum, OrderType, TitleLen, ContentLen
    Dim ShowClassName, ShowAuthor, ShowDateType, ShowHits, ShowHotSign
    Dim ShowNewSign, ShowTips, OpenType, UrlType, ShowPropertyType
    Dim Author, Cols, CssNameA, CssName1, CssName2
    Dim rsJs, sqlJs, trs
    Dim ContentType
    If Action = "SaveAdd" Then
        ID = 0
    Else
        ID = PE_CLng(Trim(Request("ID")))
        If ID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ʧ��</li>"
            Exit Sub
        End If
    End If
    JsName = Trim(Request("JsName"))
    JsReadme = Trim(Request("JsReadme"))
    JsFileName = Trim(Request("JsFileName"))
    ShowType = PE_CLng(Trim(Request("ShowType")))
    PhotoNum = PE_CLng(Trim(Request("PhotoNum")))
    ClassID = Trim(Request("ClassID"))
    IncludeChild = PE_CLng(Trim(Request("IncludeChild")))
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    IsHot = PE_CLng(Trim(Request("IsHot")))
    IsElite = PE_CLng(Trim(Request("IsElite")))
    DateNum = PE_CLng(Trim(Request("DateNum")))
    OrderType = PE_CLng(Trim(Request("OrderType")))
    TitleLen = PE_CLng(Trim(Request("TitleLen")))
    ContentLen = PE_CLng(Trim(Request("ContentLen")))
    ShowClassName = PE_CLng(Trim(Request("ShowClassName")))
    ShowAuthor = PE_CLng(Trim(Request("ShowAuthor")))
    ShowDateType = PE_CLng(Trim(Request("ShowDateType")))
    ShowHits = PE_CLng(Trim(Request("ShowHits")))
    ShowHotSign = PE_CLng(Trim(Request("ShowHotSign")))
    ShowNewSign = PE_CLng(Trim(Request("ShowNewSign")))
    ShowTips = PE_CLng(Trim(Request("ShowTips")))
    OpenType = PE_CLng(Trim(Request("OpenType")))
    UrlType = PE_CLng(Trim(Request("UrlType")))
    ShowPropertyType = PE_CLng(Trim(Request("ShowPropertyType")))
    Author = ReplaceBadChar(Trim(Request("Author")))
    Cols = PE_CLng1(Trim(Request("Cols")))
    CssNameA = ReplaceBadChar(Trim(Request("CssNameA")))
    CssName1 = ReplaceBadChar(Trim(Request("CssName1")))
    CssName2 = ReplaceBadChar(Trim(Request("CssName2")))
    ContentType = PE_CLng(Trim(Request("ContentType")))

    If JsName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>JS�������Ʋ���Ϊ�գ�</li>"
    Else
        JsName = ReplaceBadChar(JsName)
        Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsName='" & JsName & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ָ����JS���������Ѿ����ڣ�</li>"
        End If
        Set trs = Nothing
    End If
    If JsFileName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>JS�ļ�������Ϊ�գ�</li>"
    Else
        If IsValidJsFileName(JsFileName, ContentType) = False Then
            FoundErr = True
            If ContentType = 0 Then
                ErrMsg = ErrMsg & "<li>�������˷Ƿ���JS�ļ�����</li>"
            Else
                ErrMsg = ErrMsg & "<li>�������˷Ƿ���Html�ļ�����</li>"
            End If
        Else
            Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsFileName='" & JsFileName & "'")
            If Not (trs.BOF And trs.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>ָ����JS�ļ����Ѿ����ڣ�</li>"
            End If
            Set trs = Nothing
        End If
    End If
    If PhotoNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "��Ŀ�������0��</li>"
    End If
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������Ŀ����ָ��Ϊ�ⲿ��Ŀ��</li>"
    Else
        ClassID = PE_CLng(ClassID)
    End If
    If FoundErr = True Then Exit Sub

    Config = ShowType & "|" & PhotoNum & "|" & ClassID & "|" & IncludeChild & "|" & IsHot & "|" & IsElite & "|" & DateNum & "|" & OrderType & "|" & TitleLen & "|" & ContentLen & "|" & ShowClassName & "|" & ShowAuthor & "|" & ShowDateType & "|" & ShowHits & "|" & ShowHotSign & "|" & ShowNewSign & "|" & ShowTips & "|" & OpenType & "|" & UrlType & "|" & ShowPropertyType & "|" & SpecialID & "|" & Author & "|" & Cols & "|" & CssNameA & "|" & CssName1 & "|" & CssName2

    Set rsJs = Server.CreateObject("ADODB.Recordset")
    If Action = "SaveAdd" Then
        sqlJs = "select top 1 * from PE_JsFile"
        rsJs.Open sqlJs, Conn, 1, 3
        rsJs.addnew
        rsJs("JsType") = 0
        rsJs("ChannelID") = ChannelID
    Else
        sqlJs = "select * from PE_JsFile where ID=" & ID
        rsJs.Open sqlJs, Conn, 1, 3
        If rsJs.BOF And rsJs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ����JS�ļ���</li>"
            rsJs.Close
            Set rsJs = Nothing
            Exit Sub
        End If
    End If
    rsJs("JsName") = JsName
    rsJs("JsReadme") = JsReadme
    rsJs("JsFileName") = JsFileName
    rsJs("Config") = Config
    rsJs("ContentType") = ContentType
    rsJs.Update
    rsJs.Close
    Set rsJs = Nothing
    
    Call WriteSuccessMsg("����JS�ļ����óɹ���", "Admin_PhotoJS.asp?ChannelID=" & ChannelID)
    Call CreateJS(ID)
End Sub

Sub SaveJS_Pic()
    Dim ID, JsName, JsReadme, JsFileName, Config
    Dim ClassID, SpecialID, IncludeChild, PhotoNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType
    Dim rsJs, sqlJs, trs
    Dim ContentType
    If Action = "SaveAddPic" Then
        ID = 0
    Else
        ID = PE_CLng(Trim(Request("ID")))
        If ID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ʧ��</li>"
            Exit Sub
        End If
    End If
    JsName = Trim(Request("JsName"))
    JsReadme = Trim(Request("JsReadme"))
    JsFileName = Trim(Request("JsFileName"))
    ClassID = Trim(Request("ClassID"))
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    IncludeChild = PE_CLng(Trim(Request("IncludeChild")))
    PhotoNum = PE_CLng(Trim(Request("PhotoNum")))
    IsHot = PE_CLng(Trim(Request("IsHot")))
    IsElite = PE_CLng(Trim(Request("IsElite")))
    DateNum = PE_CLng(Trim(Request("DateNum")))
    OrderType = PE_CLng(Trim(Request("OrderType")))
    ShowType = PE_CLng(Trim(Request("ShowType")))
    ImgWidth = PE_CLng(Trim(Request("ImgWidth")))
    ImgHeight = PE_CLng(Trim(Request("ImgHeight")))
    TitleLen = PE_CLng(Trim(Request("TitleLen")))
    ContentLen = PE_CLng(Trim(Request("ContentLen")))
    ShowTips = PE_CLng(Trim(Request("ShowTips")))
    Cols = PE_CLng1(Trim(Request("Cols")))
    UrlType = PE_CLng(Trim(Request("UrlType")))
    ContentType = PE_CLng(Trim(Request("ContentType")))
    
    If JsName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>JS�������Ʋ���Ϊ�գ�</li>"
    Else
        JsName = ReplaceBadChar(JsName)
        Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsName='" & JsName & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ָ����JS���������Ѿ����ڣ�</li>"
        End If
        Set trs = Nothing
    End If
    If JsFileName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>JS�ļ�������Ϊ�գ�</li>"
    Else
        If IsValidJsFileName(JsFileName, ContentType) = False Then
            FoundErr = True
            If ContentType = 0 Then
                ErrMsg = ErrMsg & "<li>�������˷Ƿ���JS�ļ�����</li>"
            Else
                ErrMsg = ErrMsg & "<li>�������˷Ƿ���Html�ļ�����</li>"
            End If
        Else
            Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsFileName='" & JsFileName & "'")
            If Not (trs.BOF And trs.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>ָ����JS�ļ����Ѿ����ڣ�</li>"
            End If
            Set trs = Nothing
        End If
    End If
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������Ŀ����ָ��Ϊ�ⲿ��Ŀ��</li>"
    Else
        ClassID = PE_CLng(ClassID)
    End If
    If PhotoNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "��Ŀ�������0��</li>"
    End If
    If FoundErr = True Then Exit Sub

    Config = ClassID & "|" & IncludeChild & "|" & PhotoNum & "|" & IsHot & "|" & IsElite & "|" & DateNum & "|" & OrderType & "|" & ShowType & "|" & ImgWidth & "|" & ImgHeight & "|" & TitleLen & "|" & ContentLen & "|" & ShowTips & "|" & Cols & "|" & UrlType & "|" & SpecialID

    Set rsJs = Server.CreateObject("ADODB.Recordset")
    If Action = "SaveAddPic" Then
        sqlJs = "select top 1 * from PE_JsFile"
        rsJs.Open sqlJs, Conn, 1, 3
        rsJs.addnew
        rsJs("JsType") = 1
        rsJs("ChannelID") = ChannelID
    Else
        sqlJs = "select * from PE_JsFile where ID=" & ID
        rsJs.Open sqlJs, Conn, 1, 3
        If rsJs.BOF And rsJs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ����JS�ļ���</li>"
            rsJs.Close
            Set rsJs = Nothing
            Exit Sub
        End If
    End If
    rsJs("JsName") = JsName
    rsJs("JsReadme") = JsReadme
    rsJs("JsFileName") = JsFileName
    rsJs("Config") = Config
    rsJs("ContentType") = ContentType
    rsJs.Update
    rsJs.Close
    Set rsJs = Nothing
    
    Call WriteSuccessMsg("����JS�ļ����óɹ���", "Admin_PhotoJS.asp?ChannelID=" & ChannelID)
    Call CreateJS(ID)
End Sub
%>
