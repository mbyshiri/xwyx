<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_JS.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub Add()
    Response.Write "<form action='Admin_PhotoJS.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>添加新的JS文件（普通列表方式）</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif("", "", 0, "")

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>显示样式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowType' id='ShowType'>"
    Response.Write "          <option value='1' selected>普通列表</option>"
    Response.Write "          <option value='2'>表格式</option>"
    Response.Write "          <option value='3'>各项独立式</option>"
    Response.Write "          <option value='4'>DIV输出</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>作者：</strong></td>"
    Response.Write "      <td height='25'><input name='Author' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>如果不为空，则只显示指定作者的" & ChannelShortName & "，用于个人" & ChannelShortName & "集。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "数目：</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='10' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属栏目：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>所有栏目</option>" & GetClass_Option(0) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' value='1' checked>包含子栏目&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>注意：</b></font>不能指定为外部栏目"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属专题：</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>不属于任何专题</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "属性：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' value='1'> 热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' value='1'> 推荐" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>如果都不选，将显示所有" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>日期范围：</strong></td>"
    Response.Write "      <td height='25'>只显示最近 <input name='DateNum' type='text' id='DateNum' value='30' size='5' maxlength='3'> 天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;如果为空或0，则显示所有天数的" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>排序方法：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option value='1' selected>" & ChannelShortName & "ID（降序）</option>"
    Response.Write "          <option value='2'>" & ChannelShortName & "ID（升序）</option>"
    Response.Write "          <option value='3'>更新时间（降序）</option>"
    Response.Write "          <option value='4'>更新时间（升序）</option>"
    Response.Write "          <option value='5'>点击次数（降序）</option>"
    Response.Write "          <option value='6'>点击次数（升序）</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>属性图片样式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowPropertyType' id='ShowPropertyType'>"
    Response.Write "          <option value='0' >不显示</option>"
    Response.Write "          <option value='2'>符号</option>"
    Response.Write "          <option value='1' selected>小图片（样式1）</option>"
    Response.Write "          <option value='3' selected>小图片（样式2）</option>"
    Response.Write "          <option value='4' selected>小图片（样式3）</option>"
    Response.Write "          <option value='5' selected>小图片（样式4）</option>"
    Response.Write "          <option value='6' selected>小图片（样式5）</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "名称字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='30' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果为0，则显示完整名称。字母算一个字符，汉字算两个字符。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "简介字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='0' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果大于0，则在" & ChannelShortName & "名称下方显示指定字数的" & ChannelShortName & "简介</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='50' align='right'><strong>显示内容：</strong></td>"
    Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowClassName' type='checkbox' id='ShowClassName' value='1'>所属栏目</td>"
    Response.Write "          <td><input name='ShowAuthor' type='checkbox' id='ShowAuthor' value='1'>作者</td>"
    Response.Write "          <td>更新时间"
    Response.Write "            <select name='ShowDateType' id='ShowDateType'>"
    Response.Write "              <option value='0'>不显示</option>"
    Response.Write "              <option value='1'>年月日</option>"
    Response.Write "              <option value='2'>月日</option>"
    Response.Write "              <option value='3'>月-日</option>"
    Response.Write "            </select>"
    Response.Write "          </td>"
    Response.Write "          <td><input name='ShowHits' type='checkbox' id='ShowHits' value='1' checked>点击次数</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowHotSign' type='checkbox' id='ShowHotSign' value='1'>热门" & ChannelShortName & "标志</td>"
    Response.Write "          <td><input name='ShowNewSign' type='checkbox' id='ShowNewSign' value='1'>最新" & ChannelShortName & "标志</td>"
    Response.Write "          <td><input name='ShowTips' type='checkbox' id='ShowTips' value='1'>显示提示信息</td>"
    Response.Write "          <td> </td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "打开方式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OpenType' id='OpenType'>"
    Response.Write "          <option value='0'>在原窗口打开</option>"
    Response.Write "          <option value='1' selected>在新窗口打开</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>链接地址选项：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option value='0' selected>使用相对路径</option>"
    Response.Write "          <option value='1'>使用包含完整网址的绝对路径</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>每行标题列数：</strong></td>"
    Response.Write "      <td height='25'><input name='Cols' type='text' value='1' size='5' maxlength='3'> <font color='#FF0000'>每行显示标题的列数</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>CSS风格类名：</strong></td>"
    Response.Write "      <td height='25'><input name='CssNameA' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>列表中文字链接调用的CSS类名</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>风格样式1：</strong></td>"
    Response.Write "      <td height='25'><input name='CssName1' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>列表中奇数行的CSS效果的类名</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>风格样式2：</strong></td>"
    Response.Write "      <td height='25'><input name='CssName2' type='text' value='' size='10' maxlength='20'> <font color='#FF0000'>列表中偶数行的CSS效果的类名</font></td>"
    Response.Write "    </tr>"

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 添 加 '>"
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
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    End If
    sqlJs = "select * from PE_JsFile where ID=" & ID
    Set rsJs = Conn.Execute(sqlJs)
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的JS文件！</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If
    JsConfig = Split(rsJs("Config") & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0", "|")

    Response.Write "<form action='Admin_PhotoJS.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>修改参数（普通列表方式）</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif(rsJs("JsName"), rsJs("JsReadme"), rsJs("ContentType"), rsJs("JsFileName"))

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>显示样式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowType' id='ShowType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 1) & ">普通列表</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 2) & ">表格式</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 3) & ">各项独立式</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(0)), 4) & ">DIV输出</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>作者：</strong></td>"
    Response.Write "      <td height='25'><input name='Author' type='text' value='" & ZeroToEmpty(JsConfig(21)) & "' size='10' maxlength='20'> <font color='#FF0000'>如果不为空，则只显示指定作者的" & ChannelShortName & "，用于个人" & ChannelShortName & "集。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "数目：</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='" & JsConfig(1) & "' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属栏目：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>所有栏目</option>" & GetClass_Option(PE_CLng(JsConfig(2))) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' " & RadioValue(PE_CLng(JsConfig(3)), 1) & ">包含子栏目&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>注意：</b></font>不能指定为外部栏目"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属专题：</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>不属于任何专题</option>" & GetSpecial_Option(PE_CLng(JsConfig(20))) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "属性：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' " & RadioValue(PE_CLng(JsConfig(4)), 1) & ">热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' " & RadioValue(PE_CLng(JsConfig(5)), 1) & ">推荐" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>如果都不选，将显示所有" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>日期范围：</strong></td>"
    Response.Write "      <td height='25'>只显示最近 <input name='DateNum' type='text' id='DateNum' value='" & JsConfig(6) & "' size='5' maxlength='3'> 天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;如果为空或0，则显示所有天数的" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>排序方法：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 1) & ">" & ChannelShortName & "ID（降序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 2) & ">" & ChannelShortName & "ID（升序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 3) & ">更新时间（降序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 4) & ">更新时间（升序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 5) & ">点击次数（降序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(7)), 6) & ">点击次数（升序）</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>属性图片样式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowPropertyType' id='ShowPropertyType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 0) & ">不显示</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 2) & ">符号</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 1) & ">小图片（样式1）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 3) & ">小图片（样式2）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 4) & ">小图片（样式3）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 5) & ">小图片（样式4）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(19)), 6) & ">小图片（样式5）</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "名称字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='" & JsConfig(8) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果为0，则显示完整名称。字母算一个字符，汉字算两个字符。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "简介字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='" & JsConfig(9) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果大于0，则在" & ChannelShortName & "名称下方显示指定字数的" & ChannelShortName & "简介</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='50' align='right'><strong>显示内容：</strong></td>"
    Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowClassName' type='checkbox' id='ShowClassName' " & RadioValue(PE_CLng(JsConfig(10)), 1) & ">所属栏目</td>"
    Response.Write "          <td><input name='ShowAuthor' type='checkbox' id='ShowAuthor' " & RadioValue(PE_CLng(JsConfig(11)), 1) & ">作者</td>"
    Response.Write "          <td>更新时间"
    Response.Write "            <select name='ShowDateType' id='ShowDateType'>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 0) & ">不显示</option>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 1) & ">年月日</option>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 2) & ">月日</option>"
    Response.Write "              <option " & OptionValue(PE_CLng(JsConfig(12)), 3) & ">月-日</option>"
    Response.Write "            </select>"
    Response.Write "          </td>"
    Response.Write "          <td><input name='ShowHits' type='checkbox' id='ShowHits' " & RadioValue(PE_CLng(JsConfig(13)), 1) & ">点击次数</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowHotSign' type='checkbox' id='ShowHotSign' " & RadioValue(PE_CLng(JsConfig(14)), 1) & ">热门" & ChannelShortName & "标志</td>"
    Response.Write "          <td><input name='ShowNewSign' type='checkbox' id='ShowNewSign' " & RadioValue(PE_CLng(JsConfig(15)), 1) & ">最新" & ChannelShortName & "标志</td>"
    Response.Write "          <td><input name='ShowTips' type='checkbox' id='ShowTips' " & RadioValue(PE_CLng(JsConfig(16)), 1) & ">显示提示信息</td>"
    Response.Write "          <td> </td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "打开方式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OpenType' id='OpenType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(17)), 0) & ">在原窗口打开</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(17)), 1) & ">在新窗口打开</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>链接地址选项：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(18)), 0) & ">使用相对路径</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(18)), 1) & ">使用包含完整网址的绝对路径</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>每行标题列数：</strong></td>"
    Response.Write "      <td height='25'><input name='Cols' type='text' value='" & PE_CLng(JsConfig(22)) & "' size='5' maxlength='3'> <font color='#FF0000'>每行显示标题的列数</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>CSS风格类名：</strong></td>"
    Response.Write "      <td height='25'><input name='CssNameA' type='text' value='" & ZeroToEmpty(JsConfig(23)) & "' size='10' maxlength='20'> <font color='#FF0000'>列表中文字链接调用的CSS类名</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>风格样式1：</strong></td>"
    Response.Write "      <td height='25'><input name='CssName1' type='text' value='" & ZeroToEmpty(JsConfig(24)) & "' size='10' maxlength='20'> <font color='#FF0000'>列表中奇数行的CSS效果的类名</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>风格样式2：</strong></td>"
    Response.Write "      <td height='25'><input name='CssName2' type='text' value='" & ZeroToEmpty(JsConfig(25)) & "' size='10' maxlength='20'> <font color='#FF0000'>列表中偶数行的CSS效果的类名</font></td>"
    Response.Write "    </tr>"

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='ID' type='hidden' id='ID' value='" & ID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='保存修改结果'>"
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
    Response.Write "      <td height='22' colspan='2' align='center'><strong>添加新的JS文件（图片列表方式）</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif("", "", 0, "")

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属栏目：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>所有栏目</option>" & GetClass_Option(0) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' value='1' checked>包含子栏目&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>注意：</b></font>不能指定为外部栏目"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属专题：</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>不属于任何专题</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "数目：</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='4' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "属性：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' value='1'> 热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' value='1'> 推荐" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>如果都不选，将显示所有" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>日期范围：</strong></td>"
    Response.Write "      <td height='25'>只显示最近 <input name='DateNum' type='text' id='DateNum' value='30' size='5' maxlength='3'> 天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;如果为空或0，则显示所有天数的" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>排序方法：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option value='1' selected>" & ChannelShortName & "ID（降序）</option>"
    Response.Write "          <option value='2'>" & ChannelShortName & "ID（升序）</option>"
    Response.Write "          <option value='3'>更新时间（降序）</option>"
    Response.Write "          <option value='4'>更新时间（升序）</option>"
    Response.Write "          <option value='5'>点击次数（降序）</option>"
    Response.Write "          <option value='6'>点击次数（升序）</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>显示样式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='ShowType' type='radio' value='1' checked> 图片+名称+内容简介：上下排列<br>"
    Response.Write "        <input name='ShowType' type='radio' value='2'> （图片+名称：上下排列）+内容简介：左右排列<br>"
    Response.Write "        <input name='ShowType' type='radio' value='3'> 图片+（名称+内容简介：上下排列）：左右排列"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><b>首页图片设置：</b></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        宽度： <input name='ImgWidth' type='text' id='ImgWidth' value='130' size='5' maxlength='3'> 像素&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        高度： <input name='ImgHeight' type='text' id='ImgHeight' value='90' size='5' maxlength='3'> 像素"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "名称字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='20' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>若为0，则不显示名称；若为-1，则显示完整名称。字母算一个字符，汉字算两个字符。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "简介字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='0' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果大于0，则显示指定字数的" & ChannelShortName & "简介</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>显示内容：</strong></td>"
    Response.Write "      <td height='25'><input name='ShowTips' type='checkbox' id='ShowTips' value='1'> 显示作者、更新时间、点击数等提示信息</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>每行显示" & ChannelShortName & "数：</strong></td>"
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
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;超过指定列数就会换行"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>链接地址选项：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option value='0' selected>使用相对路径</option>"
    Response.Write "          <option value='1'>使用包含完整网址的绝对路径</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAddPic'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 添 加 '>"
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
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    End If
    sqlJs = "select * from PE_JsFile where ID=" & ID
    Set rsJs = Conn.Execute(sqlJs)
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的JS文件！</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If
    JsConfig = Split(rsJs("Config") & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0", "|")

    Response.Write "<form action='Admin_PhotoJS.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>修改参数（图片列表方式）</strong></td>"
    Response.Write "    </tr>"

    Call JsBaseInif(rsJs("JsName"), rsJs("JsReadme"), rsJs("ContentType"), rsJs("JsFileName"))

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属栏目：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ClassID'><option value='0'>所有栏目</option>" & GetClass_Option(PE_CLng(JsConfig(0))) & "</select>"
    Response.Write "        <input type='checkbox' name='IncludeChild' " & RadioValue(PE_CLng(JsConfig(1)), 1) & ">包含子栏目&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'><b>注意：</b></font>不能指定为外部栏目"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>所属专题：</strong></td>"
    Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'><option value=''>不属于任何专题</option>" & GetSpecial_Option(PE_CLng(JsConfig(15))) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "数目：</strong></td>"
    Response.Write "      <td height='25'><input name='PhotoNum' type='text' value='" & JsConfig(2) & "' size='5' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "属性：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='IsHot' type='checkbox' id='IsHot' " & RadioValue(PE_CLng(JsConfig(3)), 1) & ">热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='IsElite' type='checkbox' id='IsElite' " & RadioValue(PE_CLng(JsConfig(4)), 1) & ">推荐" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <font color='#FF0000'>如果都不选，将显示所有" & ChannelShortName & "</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>日期范围：</strong></td>"
    Response.Write "      <td height='25'>只显示最近 <input name='DateNum' type='text' id='DateNum' value='" & JsConfig(5) & "' size='5' maxlength='3'> 天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;如果为空或0，则显示所有天数的" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>排序方法：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OrderType' id='OrderType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 1) & ">" & ChannelShortName & "ID（降序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 2) & ">" & ChannelShortName & "ID（升序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 3) & ">更新时间（降序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 4) & ">更新时间（升序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 5) & ">点击次数（降序）</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(6)), 6) & ">点击次数（升序）</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>显示样式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <input name='ShowType' type='radio' " & RadioValue(PE_CLng(JsConfig(7)), 1) & ">图片+名称+内容简介：上下排列<br>"
    Response.Write "        <input name='ShowType' type='radio' " & RadioValue(PE_CLng(JsConfig(7)), 2) & ">（图片+名称：上下排列）+内容简介：左右排列<br>"
    Response.Write "        <input name='ShowType' type='radio' " & RadioValue(PE_CLng(JsConfig(7)), 3) & ">图片+（名称+内容简介：上下排列）：左右排列"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><b>首页图片设置：</b></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        宽度： <input name='ImgWidth' type='text' id='ImgWidth' value='" & JsConfig(8) & "' size='5' maxlength='3'> 像素&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        高度： <input name='ImgHeight' type='text' id='ImgHeight' value='" & JsConfig(9) & "' size='5' maxlength='3'> 像素"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "名称字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value='" & JsConfig(10) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>若为0，则不显示名称；若为-1，则显示完整名称。字母算一个字符，汉字算两个字符。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>" & ChannelShortName & "简介字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value='" & JsConfig(11) & "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果大于0，则显示指定字数的" & ChannelShortName & "简介</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>显示内容：</strong></td>"
    Response.Write "      <td height='25'><input name='ShowTips' type='checkbox' id='ShowTips' " & RadioValue(PE_CLng(JsConfig(12)), 1) & ">显示作者、更新时间、点击数等提示信息</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>每行显示" & ChannelShortName & "数：</strong></td>"
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
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;超过指定列数就会换行"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>链接地址选项：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='UrlType' id='OpenType'>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(14)), 0) & ">使用相对路径</option>"
    Response.Write "          <option " & OptionValue(PE_CLng(JsConfig(14)), 1) & ">使用包含完整网址的绝对路径</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='ID' type='hidden' id='ID' value='" & ID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModifyPic'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='保存修改结果'>"
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
            ErrMsg = ErrMsg & "<li>参数丢失！</li>"
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
        ErrMsg = ErrMsg & "<li>JS代码名称不能为空！</li>"
    Else
        JsName = ReplaceBadChar(JsName)
        Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsName='" & JsName & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>指定的JS代码名称已经存在！</li>"
        End If
        Set trs = Nothing
    End If
    If JsFileName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>JS文件名不能为空！</li>"
    Else
        If IsValidJsFileName(JsFileName, ContentType) = False Then
            FoundErr = True
            If ContentType = 0 Then
                ErrMsg = ErrMsg & "<li>你输入了非法的JS文件名！</li>"
            Else
                ErrMsg = ErrMsg & "<li>你输入了非法的Html文件名！</li>"
            End If
        Else
            Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsFileName='" & JsFileName & "'")
            If Not (trs.BOF And trs.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>指定的JS文件名已经存在！</li>"
            End If
            Set trs = Nothing
        End If
    End If
    If PhotoNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "数目必须大于0！</li>"
    End If
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>所属栏目不能指定为外部栏目！</li>"
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
            ErrMsg = ErrMsg & "<li>找不到指定的JS文件！</li>"
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
    
    Call WriteSuccessMsg("保存JS文件设置成功！", "Admin_PhotoJS.asp?ChannelID=" & ChannelID)
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
            ErrMsg = ErrMsg & "<li>参数丢失！</li>"
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
        ErrMsg = ErrMsg & "<li>JS代码名称不能为空！</li>"
    Else
        JsName = ReplaceBadChar(JsName)
        Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsName='" & JsName & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>指定的JS代码名称已经存在！</li>"
        End If
        Set trs = Nothing
    End If
    If JsFileName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>JS文件名不能为空！</li>"
    Else
        If IsValidJsFileName(JsFileName, ContentType) = False Then
            FoundErr = True
            If ContentType = 0 Then
                ErrMsg = ErrMsg & "<li>你输入了非法的JS文件名！</li>"
            Else
                ErrMsg = ErrMsg & "<li>你输入了非法的Html文件名！</li>"
            End If
        Else
            Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID & " and ID<>" & ID & " and JsFileName='" & JsFileName & "'")
            If Not (trs.BOF And trs.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>指定的JS文件名已经存在！</li>"
            End If
            Set trs = Nothing
        End If
    End If
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>所属栏目不能指定为外部栏目！</li>"
    Else
        ClassID = PE_CLng(ClassID)
    End If
    If PhotoNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "数目必须大于0！</li>"
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
            ErrMsg = ErrMsg & "<li>找不到指定的JS文件！</li>"
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
    
    Call WriteSuccessMsg("保存JS文件设置成功！", "Admin_PhotoJS.asp?ChannelID=" & ChannelID)
    Call CreateJS(ID)
End Sub
%>
