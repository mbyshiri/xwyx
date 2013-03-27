<%@language=vbscript codepage=936 %>
<%
Option Explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="Conn_Counter.asp"-->
<!--#include file="../Include/PowerEasy.Common.Security.asp"-->
<%
Dim RegCount_Fill, IntervalNum, OnlineTime
Dim style, theurl
If IsEmpty(Application("RegFields_Fill")) Or IsEmpty(Application("IntervalNum")) Then
    Call OpenConn_Counter
    Dim rs
    Set rs = Conn_Counter.Execute("select * from PE_StatInfoList")
    If Not rs.bof And Not rs.EOF Then
        IntervalNum = rs("IntervalNum")
        RegCount_Fill = rs("RegFields_Fill")
        OnlineTime = rs("OnlineTime")
        Application("OnlineTime") = OnlineTime
        Application("IntervalNum") = IntervalNum
        Application("RegFields_Fill") = RegCount_Fill
    End If
    Set rs = Nothing
    Call CloseConn_Counter
Else
    IntervalNum = Application("IntervalNum")
    RegCount_Fill = Application("RegFields_Fill")
End If

'正则表达式相关的变量
Dim regEx, Match, Match2, Matches, Matches2
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.MultiLine = True

style = FilterJS(Request("style"))
theurl = "http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
If Right(theurl, 1) <> "/" Then
    theurl = theurl & "/"
End If
%>

var style      ='<%=style%>';
var url        ='<%=theurl%>';
var IntervalNum=<%=IntervalNum%>;
var i=0;
<%
If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
%>
PowerEasyRef(0);
<%
End If
%>
document.write("<scr"+"ipt language=javascript src="+url+"counter.asp?style="+style+"&Referer="+escape(document.referrer)+"&Timezone="+escape((new Date()).getTimezoneOffset())+"&Width="+escape(screen.width)+"&Height="+escape(screen.height)+"&Color="+escape(screen.colorDepth)+"></sc"+"ript>");
function PowerEasyRef(){
    if(i <= IntervalNum){
        var PowerEasyImg=new Image();
        PowerEasyImg.src=url+'statonline.asp';
        setTimeout('PowerEasyRef()',60000);
    }
    i+=1;
}

