<%
Const SystemDatabaseType = "ACCESS"     'ϵͳ���ݿ����ͣ�"SQL"ΪMS SQL2000���ݿ⣬"ACCESS"ΪMS ACCESS 2000���ݿ�

'�����ACCESS���ݿ⣬�������޸ĺ���������ݿ���ļ���
Const DBFileName = "\database\SiteWeaver.mdb"      'ACCESS���ݿ���ļ�������ʹ���������վ��Ŀ¼�ĵľ���·��
                                        '����ǰ�װ����վ��Ŀ¼��ֱ���޸��ļ������ɡ�����ǰ�װ����վĳһĿ¼�£�����ǰ����ϴ�Ŀ¼��
                                        '���磬ϵͳ��װ�ڡ�http://www.powereasy.net/PE2006/��Ŀ¼�£�PE2006Ϊ��װĿ¼����������Ӧ���޸�Ϊ��Const DBFileName = "\PE2006\database\SiteWeaver6.5.mdb"

'�����SQL���ݿ⣬�������޸ĺ��������ݿ�ѡ��
Const SqlUsername = "PowerEasy"           'SQL���ݿ��û���
Const SqlPassword = "PowerEasy*9988"          'SQL���ݿ��û�����
Const SqlDatabaseName = "SiteWeaver66"       'SQL���ݿ���
Const SqlHostIP = "(local)"                 'SQL����IP��ַ�����أ�ָ��վ�����ݿ���ͬһ̨�������ϣ����á�(local)����127.0.0.1�����Ǳ�����ָ��վ�����ݿ�ֱ��ڲ�ͬ�ķ������ϣ�����д���ݿ����������ʵIP��


'���´�������Ķ�
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
        Response.Write "���ݿ����ӳ�������Conn.asp�ļ��е����ݿ�������á�"
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

