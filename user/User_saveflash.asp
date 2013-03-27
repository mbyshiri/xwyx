<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'目前是保存为172 * 130格式的BMP图象

Dim act, color_name, Create_1, imgurl, SaveFileName, dirMonth
ObjInstalled_FSO = IsObjInstalled(objName_FSO)
If ObjInstalled_FSO = False Then
    Response.Write "&&SendFlag=保存 >>> NO"
    Response.End
End If

act = Trim(request("act"))
If act = "" Then
    Call Main
Else
    Call CoverColorFile
End If


Sub Main()
    If CheckUserLogined() = False Then
        Call CloseConn
        Response.Write "&&SendFlag=保存 >>> NO"
        Exit Sub
    End If
    If Len(Trim(request("rgb_color"))) < 1000 Then
        Response.Write "&&SendFlag=保存 >>> NO"
    Else
        color_name = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Session.SessionID
        Set Create_1 = fso.CreateTextFile(server.mappath("flashimg/" & color_name & ".tcolor"))
        Create_1.Write (Trim(request("rgb_color")))
        Create_1.Close

        Dim urlwords, wordnum, i, weburl
        urlwords = Split(Trim(request.ServerVariables("SCRIPT_NAME")), "/")
        wordnum = UBound(urlwords)
        For i = 1 To wordnum - 1
            weburl = weburl & "/" & urlwords(i)
        Next

        imgurl = "http://" & request.ServerVariables("SERVER_NAME") & weburl & "/"

        imgurl = imgurl & "User_saveflash.asp?act=2&color_url=" & color_name '图片远程地址。

        Dim urs
        Set urs = Conn.Execute("select UserID from PE_User where UserName='" & UserName & "'")
        If Not (urs.EOF And urs.BOF) Then
            SaveFileName = InstallDir & "Space/" & UserName & urs("UserID") & "/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
            If fso.FolderExists(server.mappath(SaveFileName)) = False Then fso.CreateFolder server.mappath(SaveFileName)
            SaveFileName = SaveFileName & Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Session.SessionID & ".bmp"
            Call SaveImg(SaveFileName, imgurl)
        End If
        Set urs = Nothing
    End If
End Sub

Sub SaveImg(FileName, strUrl)
    Dim curlpath, Retrieval
    Set Retrieval = server.CreateObject("MSXML2.ServerXMLHTTP")
    Retrieval.Open "Get", strUrl, False, "", ""
    Retrieval.Send
    If Retrieval.ReadyState = 4 Then
        Set Ads = server.CreateObject("Adodb.Stream")
        Ads.Type = 1
        Ads.Mode = 3
        Ads.Open
        Ads.Write Retrieval.ResponseBody
        Ads.SaveToFile server.mappath(FileName), 2
        ads.Close()
        Set Ads = Nothing
    End If
    Set Retrieval = Nothing

    Response.Write "&&SendFlag=" & FileName
End Sub

Sub CoverColorFile()
    Dim whichfile, head, Colortxt, i, rline, badwords
    Response.Expires = -9999
    Response.AddHeader "Pragma", "no-cache"
    Response.AddHeader "cache-ctrol", "no-cache"
    Response.ContentType = "Image/bmp"

    '输出图像文件头
    head = ChrB(66) & ChrB(77) & ChrB(118) & ChrB(250) & ChrB(1) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(172) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(130) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(1) & ChrB(0) & ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(64) & ChrB(250) & ChrB(1) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB (0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0)
  
    Response.BinaryWrite head

    whichfile = Trim(request("color_url"))
    If IsNumeric(whichfile) Then
        whichfile = "flashimg/" & whichfile & ".tcolor"

        Set Colortxt = fso.OpenTextFile(server.mappath(whichfile), 1)
        rline = Colortxt.ReadLine
        badwords = Split(rline, "|")
        Colortxt.Close

        fso.deleteFile (server.mappath(whichfile))
 
        For i = 0 To UBound(badwords)
            Response.BinaryWrite to3(badwords(i))
        Next
    End If
End Sub

Function chn10(nums)
    Dim tmp, tmpstr, i
    nums_len = Len(nums)
    For i = 1 To nums_len
        tmp = Mid(nums, i, 1)
        If IsNumeric(tmp) Then
            tmp = tmp * 16 * (16 ^ (nums_len - i - 1))
        Else
            tmp = (Asc(UCase(tmp)) - 55) * (16 ^ (nums_len - i))
        End If
        tmpstr = tmpstr + tmp
    Next
    chn10 = tmpstr
End Function

Function to3(nums)
    Dim tmp, i
    Dim myArray()
    For i = 1 To 3
        tmp = Mid(nums, i * 2 - 1, 2)
        ReDim Preserve myArray(i)
        myArray(i) = chn10(tmp)
    Next
    to3 = ChrB(myArray(3)) & ChrB(myArray(2)) & ChrB(myArray(1))
End Function
%>
