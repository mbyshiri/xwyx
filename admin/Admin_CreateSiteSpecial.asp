<!--#include file="Admin_CreateCommon.asp"-->
<!--#include file="../Include/PowerEasy.Article.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<!--#include file="../Include/PowerEasy.Photo.asp"-->
<!--#include file="../Include/PowerEasy.Product.asp"-->
<!--#include file="../Include/PowerEasy.SiteSpecial.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************


ChannelID = 0

IsAutoCreate = PE_CBool(Trim(request("IsAutoCreate")))

'On Error Resume Next
SpecialID = Trim(request("SpecialID"))
If IsValidID(SpecialID) = False Then
    SpecialID = ""
End If

tmpPageTitle = strPageTitle    '����ҳ����⵽��ʱ�����У�����Ϊ��Ŀ������ҳѭ������ʱ��ʼֵ
tmpNavPath = strNavPath

Response.Write "<html><head><title>����ȫվר��</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf

If FileExt_SiteSpecial = ".asp" Then
    Response.Write "<br>��Ϊ��վ������δ����ȫվר������HTML���ܣ����Բ�������ȫվר���б�ҳ��"
    Response.End
End If

Dim rsCreate, sql, SpecialCount
Dim tmpDir, tmpTemplateID
PageTitle = ""
CreateType = PE_CLng(Trim(request("CreateType")))
Select Case CreateType
Case 0, 1 'ѡ����ר��
    If SpecialID = "" Then
        Response.Write "<li>��ָ��Ҫ���ɵ�ר��ID</li>"
        Response.End
    End If
    If InStr(SpecialID, ",") > 0 Then
        sql = "select * from PE_Special where ChannelID=0 and SpecialID in (" & SpecialID & ") order by TemplateID,OrderID"
    Else
        sql = "select * from PE_Special where ChannelID=0 and SpecialID=" & SpecialID & ""
    End If
Case 2 '����ר��
    sql = "select * from PE_Special where ChannelID=0 order by TemplateID,OrderID"
Case Else
    Response.Write "��������"
    Response.End
End Select

tmpDir = InstallDir & "Special"
If Not fso.FolderExists(Server.MapPath(tmpDir)) Then
    fso.CreateFolder Server.MapPath(tmpDir)
End If

tmpTemplateID = 0
SpecialCount = 0
iCount = 0
strTemplate = GetTemplate(0, 30, 0)
Set rsCreate = Server.CreateObject("ADODB.Recordset")
rsCreate.Open sql, Conn, 1, 1
If rsCreate.Bof And rsCreate.EOF Then
    TotalCreate = 0
    rsCreate.Close
    Set rsCreate = Nothing
    Response.End
Else
    TotalCreate = rsCreate.RecordCount
End If

Call MoveRecord(rsCreate)
Call ShowTotalCreate("��ר��")
Do While Not rsCreate.EOF
    ChannelID = 0
    PageTitle = ""
    If rsCreate("TemplateID") <> tmpTemplateID Then
        strTemplate = GetTemplate(0, 30, rsCreate("TemplateID"))
        tmpTemplateID = rsCreate("TemplateID")
    End If
    CurrentPage = 1
    SpecialID = rsCreate("SpecialID")

    strFileName = InstallDir & "ShowSpecial.asp?ClassID=" & ClassID & "&SpecialID=" & SpecialID
    Call GetSpecial
    MaxPerPage = MaxPerPage_Special
    tmpDir = InstallDir & "Special/" & SpecialDir
    If Not fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.CreateFolder Server.MapPath(tmpDir)
    End If
    
    tmpFileName = tmpDir & "/Index" & FileExt_SiteSpecial
    strHTML = strTemplate
    Call GetHtml_Special
    Call WriteToFile(tmpFileName, strHTML)

    iCount = iCount + 1
    SpecialCount = SpecialCount + 1
    Response.Write "<li>�ɹ����ɵ� <font color='red'><b>" & SpecialCount & " </b></font>��ר���ҳ�棺" & tmpFileName & "</li><br>" & vbCrLf
    Response.Flush


    If totalPut Mod MaxPerPage = 0 Then
        Pages = totalPut \ MaxPerPage
    Else
        Pages = totalPut \ MaxPerPage + 1
    End If
    If Pages > 1 Then
        For CurrentPage = 2 To Pages

            ChannelID = 0
            tmpFileName = tmpDir & "/List_" & Pages - CurrentPage + 1 & FileExt_SiteSpecial 'FileExt_List
            If IsAutoCreate = True And CurrentPage > 3 Then
                Call Update_ShowPage(tmpFileName, "UpdateSiteSpecial")
            Else
                strHTML = strTemplate
                Call GetHtml_Special
                Call WriteToFile(tmpFileName, strHTML)
                iCount = iCount + 1
                Response.Write "&nbsp;&nbsp;&nbsp;�ɹ����ɵ� <font color='red'><b>" & SpecialCount & " </b></font>��ר��ĵ� <font color='blue'>" & CurrentPage & "</font> ҳ�б�" & tmpFileName & "<br>" & vbCrLf
                Response.Flush
            End If
        Next
    End If
    rsCreate.MoveNext
    If iCount Mod MaxPerPage_Create = 0 Then Exit Do
Loop
rsCreate.Close
Set rsCreate = Nothing

strFileName = "Admin_CreateSiteSpecial.asp?Action=" & Action & "&CreateType=" & CreateType & "&SpecialID=" & Trim(Request("SpecialID")) & "&CreatePage=" & CurrentCreatePage + 1
strFileName = Replace(strFileName, " ", "")
If CurrentCreatePage < iTotalPage Then
    If SleepTime > 0 Then
        Response.Write "<p align='center'>" & SleepTime & "����Զ�����������һҳ��</p>" & vbCrLf
    End If
    Call Refresh(strFileName,SleepTime)		
   ' Response.Write "<meta http-equiv=""refresh"" content=" & SleepTime & ";url='" & strFileName & "'>" & vbCrLf
Else
    Response.Write "<p align='center'>�Ѿ���������ҳ�棡</p>" & vbCrLf
    If Trim(Request("ShowBack")) <> "No" Then
        Response.Write "<p align='center'><a href='Admin_CreateHTML.asp?Action=SiteSpecial'>�����ء�</a></p>" & vbCrLf
    End If
End If
Response.Write "</body></html>"
Call CloseConn


%>
