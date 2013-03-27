<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************


'******************************************************
'通行接口开关：API_Enable = True(启用) 或者 False(禁用)
'安 全 密 钥 ：API_Key 用户自定义的字符串，整合系统中所
'　　　　　　　有程序的密钥必须一致。
'远程系统配置：每个远程系统均包含两个部分，第一部分是该
'　　　　　　　系统的名称，第二部分为接口文件的URL；名称
'　　　　　　　和URL之间用"@@"分隔，多个远程系统之间用
'　　　　　　　"|"分隔。
'超 时 设 置 ：超时时间用于远程请求，这里的超时时间只是
'　　　　　　　一个基数，并非实际等待时间。默认设置为10
'　　　　　　　秒，表示DNS解析和建立连接超时时间10秒、
'　　　　　　　发送和接收数据超时时间为20秒。用户可以根
'　　　　　　　据自己的情况设定。通常在同一服务器可以设
'　　　　　　　置短一些，跨域名跨服务器设置长一些。
'******************************************************

Const API_Enable = False
Const API_Key = "API_TEST"
Const API_Urls = "博客@@http://Localhost/oblog4/api/API_Response.asp|论坛@@http://Localhost/bbs/dv_dpo.asp"      
Const API_Timeout = 10000

'以下请勿修改
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