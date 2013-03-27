<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim Conn_Counter, ConnStr_Count, db_counter
Dim PECount_Now
Dim CountSqlUsername, CountSqlPassword, CountSqlDatabaseName, CountSqlHostIP
Const CountDatabaseType = "ACCESS"  'ϵͳ���ݿ����ͣ�"SQL"ΪMS SQL2000���ݿ⣬"ACCESS"ΪMS ACCESS 2000���ݿ�
db_counter = "../count/counter.mdb"    '����ͳ�����ݿ��ļ���λ��

'�����SQL���ݿ⣬�������޸ĺ��������ݿ�ѡ��
CountSqlUsername = "PowerEasy"           'SQL���ݿ��û���
CountSqlPassword = "PowerEasy"       'SQL���ݿ��û�����
CountSqlDatabaseName = "Counter"    'SQL���ݿ���
CountSqlHostIP = "127.0.0.1"           'SQL����IP��ַ�����ؿ��á�127.0.0.1����(local)�����Ǳ���������ʵIP��


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
        Response.Write "���ݿ����ӳ�������Conn_Counter.asp�ļ��е����ݿ�������á�"
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
