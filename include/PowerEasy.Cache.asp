<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim PE_Cache
Set PE_Cache = New Cache

'**************************************************
'方法名：ClearSiteCache
'作  用：清除系统某频道缓存 参数 0 为 全部缓存
'参  数：iChannelID ---- 频道参数参数
'**************************************************
Sub ClearSiteCache(iChannelID)
    If iChannelID = 0 Then
        PE_Cache.DelAllCache
    Else
        PE_Cache.DelChannelCache (iChannelID)
    End If
End Sub

Class Cache

'对象的声明

Public ReloadTime    ' 过期时间（单位为分钟）
Public CacheName     '缓存组的名称（预留功能，当一个站点中有多个缓存组时，则需要为每个缓存组设置不同的名称）。
Private CacheData

Private Sub Class_Initialize()
    ReloadTime = 10
    CacheName = "PowerEasy"
End Sub

Private Sub Class_Terminate()

End Sub

'************************************************************
'函数名：SetValue
'作  用：设置缓存对象的值
'参  数：MyCacheName ---- 缓存对象的名称
'      vNewValue ----- 要给缓存对象的值
'返回值：True ---- 设置成功，False ---- 设置失败
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
'函数名：GetValue
'作  用：得到缓存对象的值
'参  数：MyCacheName ---- 缓存对象的名称
'返回值：缓存对象的值
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
'函数名：CacheIsEmpty
'作  用：判断当前缓存是否过期
'参  数：MyCacheName ---- 缓存对象的名称
'返回值：True ---- 已经过期，False ---- 没有过期
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
'过程名：DelCache
'作  用：手工删除一个缓存对象
'参  数：MyCacheName ---- 缓存对象的名称
'************************************************************
Public Sub DelCache(MyCacheName)
    Application.Lock
    Application.Contents.Remove (CacheName & "_" & MyCacheName)
    Application.UnLock
End Sub

'************************************************************
'过程名：DelAllCache
'作  用：删除全部缓存对象
'参  数：无
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
'过程名：DelChannelCache
'作  用：删除指定频道的缓存对象
'参  数：ChannelID ---- 频道ID
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
