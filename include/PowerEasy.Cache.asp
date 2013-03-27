<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim PE_Cache
Set PE_Cache = New Cache

'**************************************************
'��������ClearSiteCache
'��  �ã����ϵͳĳƵ������ ���� 0 Ϊ ȫ������
'��  ����iChannelID ---- Ƶ����������
'**************************************************
Sub ClearSiteCache(iChannelID)
    If iChannelID = 0 Then
        PE_Cache.DelAllCache
    Else
        PE_Cache.DelChannelCache (iChannelID)
    End If
End Sub

Class Cache

'���������

Public ReloadTime    ' ����ʱ�䣨��λΪ���ӣ�
Public CacheName     '����������ƣ�Ԥ�����ܣ���һ��վ�����ж��������ʱ������ҪΪÿ�����������ò�ͬ�����ƣ���
Private CacheData

Private Sub Class_Initialize()
    ReloadTime = 10
    CacheName = "PowerEasy"
End Sub

Private Sub Class_Terminate()

End Sub

'************************************************************
'��������SetValue
'��  �ã����û�������ֵ
'��  ����MyCacheName ---- ������������
'      vNewValue ----- Ҫ����������ֵ
'����ֵ��True ---- ���óɹ���False ---- ����ʧ��
'************************************************************
Public Function SetValue(MyCacheName, vNewValue)
    If MyCacheName <> "" Then
        CacheData = Application(CacheName & "_" & MyCacheName)
        If IsArray(CacheData) Then
            CacheData(0) = vNewValue
            CacheData(1) = Now()
        Else
            ReDim CacheData(2)
            CacheData(0) = vNewValue
            CacheData(1) = Now()
        End If
        Application.Lock
        Application(CacheName & "_" & MyCacheName) = CacheData
        Application.UnLock
        SetValue = True
    Else
        SetValue = False
    End If
End Function

'************************************************************
'��������GetValue
'��  �ã��õ���������ֵ
'��  ����MyCacheName ---- ������������
'����ֵ����������ֵ
'************************************************************
Public Function GetValue(MyChacheName)
    If MyChacheName <> "" Then
        CacheData = Application(CacheName & "_" & MyChacheName)
        If IsArray(CacheData) Then
            GetValue = CacheData(0)
        Else
            GetValue = ""
        End If
    Else
        GetValue = ""
    End If
End Function

'************************************************************
'��������CacheIsEmpty
'��  �ã��жϵ�ǰ�����Ƿ����
'��  ����MyCacheName ---- ������������
'����ֵ��True ---- �Ѿ����ڣ�False ---- û�й���
'************************************************************
Public Function CacheIsEmpty(MyCacheName)
    CacheIsEmpty = True
    CacheData = Application(CacheName & "_" & MyCacheName)
    If Not IsArray(CacheData) Then Exit Function
    If Not IsDate(CacheData(1)) Then Exit Function
    If DateDiff("s", CDate(CacheData(1)), Now()) < 60 * ReloadTime Then
        CacheIsEmpty = False
    End If
End Function

'************************************************************
'��������DelCache
'��  �ã��ֹ�ɾ��һ���������
'��  ����MyCacheName ---- ������������
'************************************************************
Public Sub DelCache(MyCacheName)
    Application.Lock
    Application.Contents.Remove (CacheName & "_" & MyCacheName)
    Application.UnLock
End Sub

'************************************************************
'��������DelAllCache
'��  �ã�ɾ��ȫ���������
'��  ������
'************************************************************
Public Sub DelAllCache()
    Dim Cacheobj, strAllCache, CacheList, i
    For Each Cacheobj In Application.Contents
        If CStr(Left(Cacheobj, Len(CacheName) + 1)) = CStr(CacheName & "_") Then
            strAllCache = strAllCache & Cacheobj & ","
        End If
    Next
    CacheList = Split(strAllCache, ",")
    If UBound(CacheList) > 0 Then
        For i = 0 To UBound(CacheList)
            Application.Lock
            Application.Contents.Remove CacheList(i)
            Application.UnLock
        Next
    End If
End Sub

'************************************************************
'��������DelChannelCache
'��  �ã�ɾ��ָ��Ƶ���Ļ������
'��  ����ChannelID ---- Ƶ��ID
'************************************************************
Public Sub DelChannelCache(ChannelID)
    Dim Cacheobj, strChannelCache, CacheList, i
    regEx.Pattern = "^" & CacheName & "_" & ChannelID & "_"
    For Each Cacheobj In Application.Contents
        If regEx.Test(Cacheobj) = True Then
            strChannelCache = strChannelCache & Cacheobj & ","
        End If
    Next
    CacheList = Split(strChannelCache, ",")
    If UBound(CacheList) > 0 Then
        For i = 0 To UBound(CacheList)
            Application.Lock
            Application.Contents.Remove CacheList(i)
            Application.UnLock
        Next
    End If
End Sub

End Class
%>
