<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Class.asp"-->
<!--#include file="../Include/PowerEasy.Special.asp"-->
<!--#include file="../Include/PowerEasy.Article.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<!--#include file="../Include/PowerEasy.Photo.asp"-->
<!--#include file="../Include/PowerEasy.Product.asp"-->
<!--#include file="../Include/PowerEasy.SiteIndex.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

ChannelID = 0
Response.Write "<html><head><title>������վ��ҳ</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf

If FileName_SiteIndex = "Index.asp" Then
    Response.Write "��Ϊ��վ������δ������վ��ҳ����HTML���ܣ����Բ���������ҳ��"
    Response.End
End If
Response.Write "����������վ��ҳ��" & InstallDir & FileName_SiteIndex & "������"

Call GetHTML_SiteIndex

Call WriteToFile(InstallDir & FileName_SiteIndex, strHTML)
Response.Write "����������������������վ��ҳ�ɹ���" & "</b>" & vbCrLf
Response.Write "</body></html>"
Call CloseConn
%>
