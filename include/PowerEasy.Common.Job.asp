<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Function GetEducation(Education)
    Dim strEducation
    Select Case Education
    Case 0
         strEducation = "没有选择学历"
    Case 1
         strEducation = "小学"
    Case 2
         strEducation = "初中"
    Case 3
         strEducation = "高中"
    Case 4
         strEducation = "中专"
    Case 5
         strEducation = "大专"
    Case 6
         strEducation = "本科"
    Case 7
         strEducation = "硕士"
    Case 8
         strEducation = "博士"
    Case 9
         strEducation = "博士后"
    Case 10
         strEducation = "其它"
    Case Else
         strEducation = "其它"
    End Select
    GetEducation = strEducation
End Function



Function GetMarital(Marital)
    Dim strMarital
    Select Case Marital
    Case 0
         strMarital = "没选择婚姻状况"
    Case 1
         strMarital = "未婚"
    Case 2
         strMarital = "已婚"
    Case 3
         strMarital = "离异"
    Case 4
         strMarital = "其它"
    Case Else
         strMarital = "其它"
    End Select
    GetMarital = strMarital
End Function


Function GetSex(Sex)
    Dim strSex
    Select Case Sex
    Case 0
         strSex = "没有选择性别"
    Case 1
         strSex = "男"
    Case 2
         strSex = "女"
    Case 3
         strSex = "保密"
    Case Else
         strSex = "保密"
    End Select
    GetSex = strSex
End Function

Function GetForeignLanguageKind(ForeignLanguageKind)
    Dim strForeignLanguageKind
    Select Case ForeignLanguageKind
    Case 0
         strForeignLanguageKind = "没选择外语语种"
    Case 1
         strForeignLanguageKind = "英语"
    Case 2
         strForeignLanguageKind = "日语"
    Case 3
         strForeignLanguageKind = "德语"
    Case 4
         strForeignLanguageKind = "法语"
    Case 5
         strForeignLanguageKind = "西班牙语"
    Case Else
         strForeignLanguageKind = "英语"
    End Select
    GetForeignLanguageKind = strForeignLanguageKind
End Function

Function GetCheckStatus(CheckStatus)
    Dim strCheckStatus
    Select Case CheckStatus
    Case 0
         strCheckStatus = "还未查看"
    Case 1
         strCheckStatus = "<font color='#2C9EC9'>选为初试</font>"
    Case 2
         strCheckStatus = "<font color='red'>通过复试一</font>"
    Case 3
         strCheckStatus = "<font color='#B51523'>通过复试二</font>"
    Case 4
         strCheckStatus = "<font color='blue'>已经录用</font>"
    Case Else
         strCheckStatus = "<font color='#00000F'>已经收到</font>"

    End Select
    GetCheckStatus = strCheckStatus
End Function

%>
