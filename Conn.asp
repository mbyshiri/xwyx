<%
Const SystemDatabaseType = "ACCESS"     '系统数据库类型，"SQL"为MS SQL2000数据库，"ACCESS"为MS ACCESS 2000数据库

'如果是ACCESS数据库，请认真修改好下面的数据库的文件名
Const DBFileName = "\database\SiteWeaver.mdb"      'ACCESS数据库的文件名，请使用相对于网站根目录的的绝对路径
                                        '如果是安装在网站根目录，直接修改文件名即可。如果是安装在网站某一目录下，则在前面加上此目录，
                                        '例如，系统安装在“http://www.powereasy.net/PE2006/”目录下（PE2006为安装目录），则这里应该修改为：Const DBFileName = "\PE2006\database\SiteWeaver6.5.mdb"

'如果是SQL数据库，请认真修改好以下数据库选项
Const SqlUsername = "PowerEasy"           'SQL数据库用户名
Const SqlPassword = "PowerEasy*9988"          'SQL数据库用户密码
Const SqlDatabaseName = "SiteWeaver66"       'SQL数据库名
Const SqlHostIP = "(local)"                 'SQL主机IP地址。本地（指网站与数据库在同一台服务器上）可用“(local)”或“127.0.0.1”，非本机（指网站与数据库分别在不同的服务器上）请填写数据库服务器的真实IP）


'以下代码请勿改动
Dim Conn
Dim PE_True, PE_False, PE_Now, PE_OrderType, PE_DatePart_D, PE_DatePart_Y, PE_DatePart_M, PE_DatePart_W, PE_DatePart_H
Sub OpenConn()
    'On Error Resume Next
    Dim ConnStr
    If SystemDatabaseType = "SQL" Then
        ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlHostIP & ";"
    Else
        ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBFileName)
    End If
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.open ConnStr
    If Err Then
        Err.Clear
        Set Conn = Nothing
        Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"
        Response.End
    End If
    If SystemDatabaseType = "SQL" Then
        PE_True = "1"
        PE_False = "0"
        PE_Now = "GetDate()"
        PE_OrderType = " desc"
        PE_DatePart_D = "d"
        PE_DatePart_Y = "yyyy"
        PE_DatePart_M = "m"
        PE_DatePart_W = "ww"
        PE_DatePart_H = "hh"
    Else
        PE_True = "True"
        PE_False = "False"
        PE_Now = "Now()"
        PE_OrderType = " asc"
        PE_DatePart_D = "'d'"
        PE_DatePart_Y = "'yyyy'"
        PE_DatePart_M = "'m'"
        PE_DatePart_W = "'ww'"
        PE_DatePart_H = "'h'"
    End If
End Sub

Sub CloseConn()
    On Error Resume Next
    If IsObject(Conn) Then
        Conn.Close
        Set Conn = Nothing
    End If
    Set regEx = Nothing
    Set PE_Cache = Nothing
End Sub
%>

