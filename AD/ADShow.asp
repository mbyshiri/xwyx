<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim ADID
ADID = PE_Clng(Trim(Request("ADID")))
If ADID > 0 Then
    Dim rsAD
    Set rsAD = Conn.Execute("select * from PE_Advertisement where ADID=" & ADID)
    If Not (rsAD.bof And rsAD.EOF) Then
        Select Case rsAD("ADType")
        Case 1
            Dim ADLinkUrl
            If rsAD("LinkUrl") <> "" Then
                If rsAD("CountClick") = True Then
                    ADLinkUrl = InstallDir & ADDir & "/ADCount.asp?Action=Click&ADID=" & rsAD("ADID")
                Else
                    ADLinkUrl = rsAD("LinkUrl")
                End If
                Response.Write "<a href='" & ADLinkUrl & "'"
                If rsAD("LinkTarget") = 0 Then
                    Response.Write " target='_self'"
                Else
                    Response.Write " target='_blank'"
                End If
                If rsAD("LinkAlt") <> "" Then
                    Response.Write " title='" & rsAD("LinkAlt") & "'"
                End If
                Response.Write ">"
            End If
            Response.Write "<img name='AD_" & rsAD("ADID") & "' id='AD_" & rsAD("ADID") & "' src='" & rsAD("ImgUrl") & "'"
            If rsAD("ImgWidth") <> 0 Then
                Response.Write " width='" & rsAD("ImgWidth") & "'"
            End If
            If rsAD("ImgHeight") <> 0 Then
                Response.Write " height='" & rsAD("ImgHeight") & "'"
            End If
            Response.Write " border='0'>"
            If rsAD("LinkUrl") <> "" Then
                Response.Write "</a>"
            End If
        Case 2
            Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0'"
            Response.Write " name='AD_" & rsAD("ADID") & "' id='AD_" & rsAD("ADID") & "'"
            Response.Write " width='" & rsAD("ImgWidth") & "'"
            Response.Write " height='" & rsAD("ImgHeight") & "'"
            Response.Write ">"
            Response.Write "<param name='movie' value='" & rsAD("ImgUrl") & "'>"
            If rsAD("FlashWmode") = 1 Then Response.Write "<param name='wmode' value='Transparent'>"
            Response.Write "<param name='quality' value='autohigh'>"
            Response.Write "<embed"
            Response.Write " name='AD_" & rsAD("ADID") & "' id='AD_" & rsAD("ADID") & "'"
            Response.Write " width='" & rsAD("ImgWidth") & "'"
            Response.Write " height='" & rsAD("ImgHeight") & "'"
            Response.Write " src='" & rsAD("ImgUrl") & "'"
            If rsAD("FlashWmode") = 1 Then Response.Write " wmode='Transparent'"
            Response.Write " quality='autohigh'"
            Response.Write " pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>"
            Response.Write "</object>"
        Case 3, 4
            Response.Write rsAD("ADIntro")
        Case 5
            Response.Write "<iframe id='AD_" & rsAD("ADID") & "' marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=no width=100% height=100% src='" & rsAD("ImgUrl") & "'>AD</iframe>"
        End Select
    End If
    rsAD.Close
    Set rsAD = Nothing
End If
Call CloseConn
%>