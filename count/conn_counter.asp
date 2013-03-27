<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim Conn_Counter, ConnStr_Count, db_counter
Dim PECount_Now
Dim CountSqlUsername, CountSqlPassword, CountSqlDatabaseName, CountSqlHostIP
Const CountDatabaseType = "ACCESS"  '系统数据库类型，"SQL"为MS SQL2000数据库，"ACCESS"为MS ACCESS 2000数据库
db_counter = "../count/counter.mdb"    '访问统计数据库文件的位置

'如果是SQL数据库，请认真修改好以下数据库选项
CountSqlUsername = "PowerEasy"           'SQL数据库用户名
CountSqlPassword = "PowerEasy"       'SQL数据库用户密码
CountSqlDatabaseName = "Counter"    'SQL数据库名
CountSqlHostIP = "127.0.0.1"           'SQL主机IP地址（本地可用“127.0.0.1”或“(local)”，非本机请用真实IP）


Sub OpenConn_Counter()
    On Error Resume Next
    If CountDatabaseType = "SQL" Then
        ConnStr_Count = "Provider = Sqloledb; User ID = " & CountSqlUsername & "; Password = " & CountSqlPassword & "; Initial Catalog = " & CountSqlDatabaseName & "; Data Source = " & CountSqlHostIP & ";"
    Else
        ConnStr_Count = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db_counter)
    End If

    Set Conn_Counter = Server.CreateObject("ADODB.Connection")
    Conn_Counter.Open ConnStr_Count
    If Err Then
        Err.Clear
        Set Conn_Counter = Nothing
        Response.Write "数据库连接出错，请检查Conn_Counter.asp文件中的数据库参数设置。"
        Response.End
    End If
    If CountDatabaseType = "SQL" Then
        PECount_Now = "getdate()"
    Else
        PECount_Now = "Now()"
    End If
End Sub

Sub CloseConn_Counter()
    On Error Resume Next
    If IsObject(Conn_Counter) Then
        Conn_Counter.Close
        Set Conn_Counter = Nothing
    End If
End Sub

Function FoundInArr(strArr, strItem, strSplit)
    Dim arrTemp, i
    FoundInArr = False
    If InStr(strArr, strSplit) > 0 Then
        arrTemp = Split(strArr, strSplit)
        For i = 0 To UBound(arrTemp)
            If Trim(arrTemp(i)) = Trim(strItem) Then
                FoundInArr = True
                Exit For
            End If
        Next
    Else
        If Trim(strArr) = Trim(strItem) Then
            FoundInArr = True
        End If
    End If
End Function

Function finddir(filepath)
    Dim i, abc
    finddir = ""
    For i = 1 To Len(filepath)
        If Left(Right(filepath, i), 1) = "/" Or Left(Right(filepath, i), 1) = "\" Then
            abc = i
            Exit For
        End If
    Next
    If abc <> 1 Then
        finddir = Left(filepath, Len(filepath) - abc + 1)
    End If
End Function

%>
