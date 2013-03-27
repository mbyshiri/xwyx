<!--#include file="../conn.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim Action, ADID
Action = Trim(Request("Action"))
ADID = Trim(Request("ADID"))
If ADID <> "" And Isnumeric(ADID) Then
    ADID = CLng(ADID)
    Call OpenConn
	If Action = "Click" Then
        Conn.Execute ("update PE_Advertisement set Clicks=Clicks+1 where ADID=" & ADID)
        Dim rsAD, ADLinkUrl
        Set rsAD = Conn.Execute("select LinkUrl from PE_Advertisement where ADID=" & ADID)
        If Not (rsAD.bof And rsAD.EOF) Then
            ADLinkUrl = rsAD("LinkUrl")
        End If
        rsAD.Close
        Set rsAD = Nothing
        If ADLinkUrl <> "" Then Response.Redirect (ADLinkUrl)
    Else
        Conn.Execute ("update PE_Advertisement set Views=Views+1 where ADID=" & ADID)
    End If
	Call CloseConn
End If
%>