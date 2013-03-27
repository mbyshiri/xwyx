<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
Const NeedCheckComeUrl = True


Dim TruePath, theFolder, theSubFolder, theFile, thisfile, FileCount, TotalSize, TotalSize_Page
Dim TotalUnit, strTotalUnit, PageUnit, strPageUnit
Dim StrFileType, strFiles
Dim strDirName, tUploadDir, ShowFileStyle
Dim RootDir
Dim strPath, strPath2, strPath3

'排序增加变量
Dim SysRootDir
Dim UpFilesPath, FS, FolderObj, SubFolderObj, FolderItem, UpLoadNumber, TempUpLoadImgSrc

'获取频道相关数据
tUploadDir = Trim(Request("UploadDir"))
If ChannelID > 0 Then

Else
    If tUploadDir = "UploadAdPic" Then
        ChannelName = "网站广告"
        UploadDir = "UploadAdPic"
        ChannelDir = ADDir
    End If
End If

'检查管理员操作权限
If AdminPurview > 1 Then
    If ChannelID > 0 Then
        If AdminPurview_Channel = "" Then
            AdminPurview_Channel = 5
        Else
            AdminPurview_Channel = PE_CLng(AdminPurview_Channel)
        End If
        If AdminPurview_Channel > 1 Then
            PurviewPassed = False
        Else
            PurviewPassed = True
        End If
    Else
        If tUploadDir = "UploadAdPic" Then
            PurviewPassed = CheckPurview_Other(AdminPurview_Others, "AD")
        Else
            PurviewPassed = False
        End If
    End If
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Response.End
    End If
End If

Dim rs, sql
Select Case ModuleType
Case 1
    strDirName = ChannelName & "的上传文件"
    sql = "select UploadFiles from PE_Article where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
Case 2
    If tUploadDir = "UploadSoftPic" Then
        UploadDir = "UploadSoftPic"
        strDirName = ChannelName & "的软件图片"
        sql = "select SoftPicUrl from PE_Soft where ChannelID=" & ChannelID
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
    Else
        strDirName = ChannelName & "的上传软件"
        sql = "select DownloadUrl from PE_Soft where ChannelID=" & ChannelID
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "$$$" & rs(0)
            End If
            rs.MoveNext
        Loop
    End If
Case 3
    strDirName = ChannelName & "的上传图片"
    sql = "select PhotoThumb,PhotoUrl from PE_Photo"
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "$$$" & rs(0)
        End If
        If rs(1) <> "" Then
            strFiles = strFiles & "$$$" & rs(1)
        End If
        rs.MoveNext
    Loop
Case 5
    strDirName = ChannelName & "的上传图片"
    sql = "select UploadFiles from PE_Product where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
Case 6
    strDirName = ChannelName & "的上传图片"
    sql = "select SupplyPicUrl from PE_Supply where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
Case 7 '清除房产模块的图片
    Dim i, HouseTable
    strDirName = ChannelName & "的上传图片"
    For i = 1 To 5
        Select Case i
        Case 1
            HouseTable = "PE_HouseCS"
        Case 2
            HouseTable = "PE_HouseCZ"
        Case 3
            HouseTable = "PE_HouseQG"
        Case 4
            HouseTable = "PE_HouseQZ"
        Case 5
            HouseTable = "PE_HouseHZ"
        End Select
        sql = "select UploadPhotos from " & HouseTable & " where ChannelID=" & ChannelID
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
    Next
Case 8 '清除人才招聘模块的图片
    strDirName = ChannelName & "的上传图片"
    sql = "select Photo from PE_Resume"
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
Case Else
    If tUploadDir = "UploadAdPic" Then
        strDirName = "上传的广告图片"
        sql = "select ImgUrl from PE_Advertisement"
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
    End If
End Select
rs.Close
Set rs = Nothing
strFiles = LCase(strFiles)

RootDir = InstallDir & ChannelDir & "/" & UploadDir

SysRootDir = RootDir & "/"
UploadDir = Trim(Request("UploadDir"))


If SysRootDir <> "" Then
    UpFilesPath = SysRootDir
Else
    UpFilesPath = "/"
End If
UpLoadNumber = 1
Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<title>管理导航菜单</title>" & vbCrLf
    
Response.Write "    <STYLE type=text/css>" & vbCrLf
Response.Write "    body  { background:#F2f2f2; margin:0px; font:9pt 宋体; FONT-SIZE: 9pt;text-decoration: none;" & vbCrLf
Response.Write "    SCROLLBAR-FACE-COLOR: #D8E5FC;" & vbCrLf
Response.Write "    SCROLLBAR-HIGHLIGHT-COLOR: #FAFAFA; SCROLLBAR-SHADOW-COLOR: #DBEBFA; SCROLLBAR-3DLIGHT-COLOR: #DBEBFA; SCROLLBAR-ARROW-COLOR: #EAF3FF; SCROLLBAR-TRACK-COLOR: #FAFAFA; SCROLLBAR-DARKSHADOW-COLOR: #FAFAFA;" & vbCrLf
Response.Write "    table  { border:0px; }" & vbCrLf
Response.Write "    td  { font:normal 12px 宋体; }" & vbCrLf
Response.Write "    img  { vertical-align:bottom; border:0px; }" & vbCrLf
Response.Write "    a  { font:normal 12px 宋体;　color:#000000; text-decoration: none;}" & vbCrLf
Response.Write "    a:link {color:#000000;text-decoration: none;}" & vbCrLf
Response.Write "    a:hover  { color:#000000;}" & vbCrLf
Response.Write "    </STYLE> " & vbCrLf
Response.Write "    <script language=""JavaScript"">" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "    function ClickClassImg(ClickObj,ClassID)" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "        var ImgSrc=ClickObj.src,OpenTF;" & vbCrLf
Response.Write "        var FolderObj=ClickObj.parentElement.children(ClickObj.parentElement.children.length-1);" & vbCrLf
Response.Write "        if (ImgSrc.indexOf('Close.gif')!=-1) {ClickObj.src='Images/Folder/Open.gif';OpenTF=true}" & vbCrLf
Response.Write "        if (ImgSrc.indexOf('EndClose.gif')!=-1) {ClickObj.src='Images/Folder/EndOpen.gif';OpenTF=true};" & vbCrLf
Response.Write "        if (ImgSrc.indexOf('Open.gif')!=-1) {ClickObj.src='Images/Folder/Close.gif';OpenTF=false;}" & vbCrLf
Response.Write "        if (ImgSrc.indexOf('EndOpen.gif')!=-1) {ClickObj.src='Images/Folder/EndClose.gif';OpenTF=false;}" & vbCrLf
Response.Write "        if (OpenTF) " & vbCrLf
Response.Write "        {" & vbCrLf
Response.Write "            if (FolderObj.src.indexOf('folderclosed.gif')!=-1) FolderObj.src='Images/Folder/folderopen.gif';" & vbCrLf
Response.Write "            ShowChildClass(ClassID);" & vbCrLf
Response.Write "        }" & vbCrLf
Response.Write "        else" & vbCrLf
Response.Write "        {" & vbCrLf
Response.Write "            if (FolderObj.src.indexOf('folderopen.gif')!=-1) FolderObj.src='Images/Folder/folderclosed.gif';" & vbCrLf
Response.Write "            HideChildClass(ClassID);" & vbCrLf
Response.Write "        }" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "    function ShowChildClass(ID)" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "        var CurrObj=null;" & vbCrLf
Response.Write "        var TRObj=document.body.getElementsByTagName('TR');" & vbCrLf
Response.Write "        for (var i=0;i<TRObj.length;i++)" & vbCrLf
Response.Write "        {" & vbCrLf
Response.Write "            CurrObj=TRObj(i);" & vbCrLf
Response.Write "            if (CurrObj.ParentID==ID)" & vbCrLf
Response.Write "            {" & vbCrLf
Response.Write "                if (CurrObj.tagName.toLowerCase()=='tr')" & vbCrLf
Response.Write "                {" & vbCrLf
Response.Write "                    CurrObj.style.display='';" & vbCrLf
Response.Write "                    ChangeImg(CurrObj,false);" & vbCrLf
Response.Write "                }" & vbCrLf
Response.Write "            }" & vbCrLf
Response.Write "        }" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "    function ChangeImg(Obj,OpenTF)" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "        var CurrObj=null,ImgSrc='';" & vbCrLf
Response.Write "        for (var i=0;i<Obj.all.length;i++)" & vbCrLf
Response.Write "        {" & vbCrLf
Response.Write "            CurrObj=Obj.all(i);" & vbCrLf
Response.Write "            if (CurrObj.tagName.toLowerCase()=='img')" & vbCrLf
Response.Write "            {" & vbCrLf
Response.Write "                ImgSrc=CurrObj.src;" & vbCrLf
Response.Write "                if (OpenTF==true)" & vbCrLf
Response.Write "                {" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('Close.gif')!=-1) CurrObj.src='Images/Folder/Open.gif';" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('EndClose.gif')!=-1) CurrObj.src='Images/Folder/EndOpen.gif';" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('Open.gif')!=-1) return;" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('EndOpen.gif')!=-1) return;" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('folderopen.gif')!=-1) return;" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('folderclosed.gif')!=-1) CurrObj.src='Images/Folder/folderopen.gif';" & vbCrLf
Response.Write "                }" & vbCrLf
Response.Write "                else" & vbCrLf
Response.Write "                {" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('Close.gif')!=-1) return;" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('EndClose.gif')!=-1) return;" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('Open.gif')!=-1) CurrObj.src='Images/Folder/Close.gif';" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('EndOpen.gif')!=-1) CurrObj.src='Images/Folder/EndClose.gif';" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('folderopen.gif')!=-1) CurrObj.src='Images/Folder/folderclosed.gif';" & vbCrLf
Response.Write "                    if (ImgSrc.indexOf('folderclosed.gif')!=-1) return;" & vbCrLf
Response.Write "                }" & vbCrLf
Response.Write "            }" & vbCrLf
Response.Write "        }" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "    function HideChildClass(ID)" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "        var CurrObj=null;" & vbCrLf
Response.Write "        var TRObj=document.body.getElementsByTagName('TR');" & vbCrLf
Response.Write "        for (var i=0;i<TRObj.length;i++)" & vbCrLf
Response.Write "        {" & vbCrLf
Response.Write "            CurrObj=TRObj(i);" & vbCrLf
Response.Write "            if (CurrObj.AllParentID!=null)" & vbCrLf
Response.Write "            {" & vbCrLf
Response.Write "                if (CurrObj.AllParentID.indexOf(ID)!=-1) CurrObj.style.display='none';" & vbCrLf
Response.Write "            }" & vbCrLf
Response.Write "        }" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "    </script>" & vbCrLf
Response.Write "" & vbCrLf

Response.Write "</head>" & vbCrLf
Response.Write "<BODY leftmargin='0' topmargin='0' marginheight='0' marginwidth='0'>" & vbCrLf
If fso.FolderExists(Server.MapPath(UpFilesPath)) Then
    Set FolderObj = fso.GetFolder(Server.MapPath(UpFilesPath))
    Set SubFolderObj = FolderObj.SubFolders
Else
    FoundErr = True
    Response.Write "上传目录不存在！"
    Response.End
End If
Response.Write "<table id='RootDir' RootDir='" & UpFilesPath & "' width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td> " & vbCrLf
Response.Write "      <table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
Response.Write "        <tr>" & vbCrLf
Response.Write "          <td colspan='2' height='28'>上传目录导航</td>" & vbCrLf
Response.Write "        </tr><tr>" & vbCrLf
Response.Write "          <td><img src='Images/Folder/folderopen.gif' width='18' height='18'></td>" & vbCrLf
Response.Write "          <td><a href='Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "' target=UploadFile_Main><span Path='" & UpFilesPath & "'  class='TempletItem'>" & UploadDir & "</span></a></td>" & vbCrLf
Response.Write "        </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "</td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "" & vbCrLf

For Each FolderItem In SubFolderObj
    If UpLoadNumber = SubFolderObj.Count Then
        TempUpLoadImgSrc = "Images/Folder/EndClose.gif"
    Else
        TempUpLoadImgSrc = "Images/Folder/Close.gif"
    End If
        Response.Write " <tr AllParentID=" & UpFilesPath & " ParentID=" & UpFilesPath & " ClassID=" & UpFilesPath & FolderItem.name & ">"
        Response.Write "     <td>"
        Response.Write " <table border='0' cellspacing='0' cellpadding='0'>"
        Response.Write "     <tr>"
        Response.Write "       <td><img Depth='1' onClick=""ClickClassImg(this,'" & UpFilesPath & FolderItem.name & "');"" src=" & TempUpLoadImgSrc & " width='16' height='22'><img src='Images/Folder/folderclosed.gif' width='18' height='18'></td>"
        Response.Write "       <td><a href='Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&CurrentDir=" & FolderItem.name & "' target=UploadFile_Main><span Path='" & UpFilesPath & FolderItem.name & "' class='TempletItem'>" & FolderItem.name & "</span</a></td>"
        Response.Write "     </tr>"
        Response.Write "   </table>"
        Response.Write " </td>"
        Response.Write " </tr>"
    If UpLoadNumber = SubFolderObj.Count Then
        Response.Write (GetChildFolderList(UpFilesPath & FolderItem.name, "", True, ""))
    Else
        Response.Write (GetChildFolderList(UpFilesPath & FolderItem.name, "", False, ""))
    End If
    UpLoadNumber = UpLoadNumber + 1
Next

Set FolderObj = Nothing
Set SubFolderObj = Nothing

Response.Write "</table>" & vbCrLf
Response.Write "</td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf


Function GetChildFolderList(FolderID, Str, EndNodeTF, TempAllParentID)
    Dim TempImageStr, ImageStr, ChildFolderNumber, AllParentID
    Dim TempSrc, TempEndNodeTF
    Dim FolderObj, SubFolderObj, FolderItem
    If EndNodeTF = True Then
        TempSrc = "<img src=""Images/Folder/blank.gif"">"
    Else
        TempSrc = "<img src=""Images/Folder/HR.gif"">"
    End If
    ChildFolderNumber = 1
    AllParentID = TempAllParentID & "," & FolderID
    ImageStr = Str & TempSrc

    FolderID = Replace(FolderID, "//", "/")
    Set FolderObj = fso.GetFolder(Server.MapPath(FolderID))
    Set SubFolderObj = FolderObj.SubFolders
    For Each FolderItem In SubFolderObj
        If ChildFolderNumber = SubFolderObj.Count Then
            TempEndNodeTF = True
            TempImageStr = "<img onClick=""ClickClassImg(this,'" & FolderID & "/" & FolderItem.name & "')"" src=""Images/Folder/EndClose.gif""><img src=""Images/Folder/folderclosed.gif"">"
        Else
            TempEndNodeTF = False
            TempImageStr = "<img onClick=""ClickClassImg(this,'" & FolderID & "/" & FolderItem.name & "')"" src=""Images/Folder/Close.gif""><img src=""Images/Folder/folderclosed.gif"">"
        End If
        GetChildFolderList = GetChildFolderList & "<tr AllParentID=""" & AllParentID & """ ParentID=""" & FolderID & """ ClassID=""" & FolderID & "/" & FolderItem.name & """ style=""display:none;""><td><table border=""0"" cellspacing=""0"" cellpadding=""0""><tr align=""left"" class=""TempletItem""><td>" & ImageStr & TempImageStr & "</td><td nowrap><a href='Admin_UploadFile_Main.asp?UploadDir=" & UploadDir & "&ChannelID=" & ChannelID & GetCurrentPath(RePath(FolderID & "/" & FolderItem.name)) & "' target=UploadFile_Main><span Path=""" & FolderID & "/" & FolderItem.name & """ class=""TempletItem"">" & FolderItem.name & "</span></a></td></tr></table></td></tr>" & Chr(13) & Chr(10)
        GetChildFolderList = GetChildFolderList & GetChildFolderList(FolderID & "/" & FolderItem.name, ImageStr, TempEndNodeTF, AllParentID)
        ChildFolderNumber = ChildFolderNumber + 1
    Next
    
    Set FolderObj = Nothing
    Set SubFolderObj = Nothing
End Function

Function RePath(S)

    RePath = Replace(S, "\", "\\")
End Function

Function GetCurrentPath(Path)
    Dim tempArr
    If Path = "" Then
        GetCurrentPath = ""
        Exit Function
    Else
        tempArr = Split(Path, "/")
        If SysRootDir <> "" Then
            GetCurrentPath = "&ParentDir=" & Replace(Replace(Path, "/" & tempArr(UBound(tempArr)), ""), "/" & tempArr(1) & "/", "") & "&CurrentDir=" & tempArr(UBound(tempArr))
        Else
            GetCurrentPath = "&ParentDir=" & Replace(Right(Path, Len(Path) - 1), "/" & tempArr(UBound(tempArr)), "") & "&CurrentDir=" & tempArr(UBound(tempArr))
        End If
    End If
End Function
%>
