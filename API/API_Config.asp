<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************


'******************************************************
'ͨ�нӿڿ��أ�API_Enable = True(����) ���� False(����)
'�� ȫ �� Կ ��API_Key �û��Զ�����ַ���������ϵͳ����
'���������������г������Կ����һ�¡�
'Զ��ϵͳ���ã�ÿ��Զ��ϵͳ�������������֣���һ�����Ǹ�
'��������������ϵͳ�����ƣ��ڶ�����Ϊ�ӿ��ļ���URL������
'����������������URL֮����"@@"�ָ������Զ��ϵͳ֮����
'��������������"|"�ָ���
'�� ʱ �� �� ����ʱʱ������Զ����������ĳ�ʱʱ��ֻ��
'��������������һ������������ʵ�ʵȴ�ʱ�䡣Ĭ������Ϊ10
'���������������룬��ʾDNS�����ͽ������ӳ�ʱʱ��10�롢
'�����������������ͺͽ������ݳ�ʱʱ��Ϊ20�롣�û����Ը�
'�����������������Լ�������趨��ͨ����ͬһ������������
'���������������ö�һЩ������������������ó�һЩ��
'******************************************************

Const API_Enable = False
Const API_Key = "API_TEST"
Const API_Urls = "����@@http://Localhost/oblog4/api/API_Response.asp|��̳@@http://Localhost/bbs/dv_dpo.asp"      
Const API_Timeout = 10000

'���������޸�
Dim arrAPIUrls, arrUrlsSP2
arrUrlsSP2 = "blank"
arrAPIUrls = Split(API_Urls,"|")
Dim tempIndex,tempAPIPath
For tempIndex = 0 To UBound(arrAPIUrls)
    tempAPIPath = Split(arrAPIUrls(tempIndex),"@@")
    arrUrlsSP2 = arrUrlsSP2 & "|" & tempAPIPath(1)
Next
arrUrlsSP2 = Replace(arrUrlsSP2,"blank|","")
arrUrlsSP2 = Split(arrUrlsSP2,"|")
%>