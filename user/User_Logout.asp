<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

If CheckUserLogined() = True Then
    Dim APISysKey
    APISysKey = LCase(MD5(UserName & API_Key,16))

    If Action <> "xml" Then
        Dim strTempMsg
        If API_Enable Then
            Dim iIndex, strLogoutParams
            strLogoutParams = "?syskey=" & APISysKey & "&username=" & UserName
            For iIndex = 0 To UBound(arrUrlsSP2)
                strTempMsg = strTempMsg & "<iframe frameborder='0' width='1' height='1' src='" & arrUrlsSP2(iIndex) & strLogoutParams &  "'></iframe>"
            Next
        End If
        strTempMsg = "您已成功注销，期待您的再次光临!" & strTempMsg
        Call WriteSuccessMsg(strTempMsg, InstallDir & "Index.asp")
    Else
        Response.Clear
        Response.ContentType = "text/xml; charset=gb2312"
        Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
        Response.Write "<body>"
        Response.Write "<user>" & UserName & "</user>"
        Response.Write "<checkstat>err</checkstat>"
        Response.Write "<errsource/>"
        Response.Write "<checkcode>0</checkcode>"
        If API_Enable Then
            Response.Write "<syskey>" & APISysKey & "</syskey>"
            Dim intIndex
            For intIndex = 0 To UBound(arrUrlsSP2)
                Response.Write "<apiurl><![CDATA[" & arrUrlsSP2(intIndex) & "]]></apiurl>"
            Next
        Else
            Response.Write "<syskey/><apiurl/>"
        End If
        Response.Write "<savecookie/>"
        Response.Write "</body>"
    End If
    Dim iSubDomainIndex, strSite_Sn
    If Enable_SubDomain Then
        For iSubDomainIndex = 0 To UBound(arrSubDomains)
            strSite_Sn = LCase(arrSubDomains(iSubDomainIndex) & Replace(Replace(DomainRoot & InstallDir, "/", ""), ".", ""))
            Response.Cookies(strSite_Sn).Domain = DomainRoot
            Response.Cookies(strSite_Sn)("UserName") = ""
            Response.Cookies(strSite_Sn)("UserPassword") = ""
            Response.Cookies(strSite_Sn)("LastPassword") = ""
        Next
    Else
        Response.Cookies(Site_Sn)("UserName") = ""
        Response.Cookies(Site_Sn)("UserPassword") = ""
        Response.Cookies(Site_Sn)("LastPassword") = ""
    End If
End If
Call CloseConn
%>