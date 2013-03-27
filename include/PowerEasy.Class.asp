<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'定义栏目设置相关的变量
Dim ClassID, ClassName, ReadMe, RootID, Depth, ParentID, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPicUrl, ItemCount, ItemID, ClassShowType
Dim EnableProtect, ClassPurview, DefaultItemSkin, DefaultItemTemplate, ItemListOrderType, ItemOpenType, Meta_Keywords_Class, Meta_Description_Class, Custom_Content_Class
Dim EnableComment, CheckComment

Sub GetClass()
    ClassName = ""
    ReadMe = ""
    Meta_Keywords_Class = ""
    Meta_Description_Class = ""
    Custom_Content_Class = ""
    ParentID = 0
    ParentPath = "0"
    Child = 0
    arrChildID = ""
    EnableProtect = 0
    MaxPerPage = 20
    DefaultItemSkin = 0
    DefaultItemTemplate = 0
    ItemListOrderType = 1
    ItemOpenType = 1
    ParentDir = "/"
    ClassDir = ""
    TemplateID = 0
    SkinID = 0
    ClassPurview = 0
    EnableComment = False
    CheckComment = False
    If ClassID > 0 Then
        Dim tClass
        Set tClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 " & " and ClassID=" & ClassID & "")
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的栏目</li>"
        Else
            ClassName = tClass("ClassName")
            ReadMe = tClass("Readme")
            Meta_Keywords_Class = tClass("Meta_Keywords")
            Meta_Description_Class = tClass("Meta_Description")
            Custom_Content_Class = tClass("Custom_Content")
            ParentID = tClass("ParentID")
            ParentPath = tClass("ParentPath")
            Child = tClass("Child")
            arrChildID = tClass("arrChildID")
            EnableProtect = tClass("EnableProtect")
            MaxPerPage = tClass("MaxPerPage")
            DefaultItemSkin = tClass("DefaultItemSkin")
            If DefaultItemSkin = 0 Then DefaultItemSkin = DefaultSkinID
            DefaultItemTemplate = tClass("DefaultItemTemplate")
            ItemListOrderType = tClass("ItemListOrderType")
            ItemOpenType = tClass("ItemOpenType")
            ParentDir = tClass("ParentDir")
            ClassDir = tClass("ClassDir")
            ClassPicUrl = tClass("ClassPicUrl")
            TemplateID = tClass("TemplateID")
            If ItemID <= 0 Then
                SkinID = tClass("SkinID")
                If SkinID = 0 Then SkinID = DefaultSkinID
            End If
            ClassPurview = tClass("ClassPurview")
            EnableComment = tClass("EnableComment")
            CheckComment = tClass("CheckComment")
            
            strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;"
            strPageTitle = strPageTitle & " >> "
            If ParentID > 0 Then
                Dim sqlPath, rsPath
                sqlPath = "select ClassID,ClassName,ParentDir,ClassDir,ClassPurview from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
                Set rsPath = Conn.Execute(sqlPath)
                Do While Not rsPath.EOF
                    '下面这一行代码是用于权限的继承。
                    If rsPath("ClassPurview") > ClassPurview Then ClassPurview = rsPath("ClassPurview")
                    strNavPath = strNavPath & "<a class='LinkPath' href='" & GetClassUrl(rsPath("ParentDir"), rsPath("ClassDir"), rsPath("ClassID"), rsPath("ClassPurview")) & "'>" & rsPath("ClassName") & "</a>&nbsp;" & strNavLink & "&nbsp;"
                    strPageTitle = strPageTitle & rsPath("ClassName") & " >> "
                    rsPath.MoveNext
                Loop
                rsPath.Close
                Set rsPath = Nothing
            End If
            strNavPath = strNavPath & "<a class='LinkPath' href='" & GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview) & "'>" & ClassName & "</a>"
            strPageTitle = strPageTitle & ClassName
        End If
        tClass.Close
        Set tClass = Nothing
    End If
End Sub

'=================================================
'函数名：GetClass_Navigation
'作  用：显示栏目导航的HTML代码
'参  数：ShowType ---- 显示样式，1为平行式，2为纵列式
'        Cols ---- 当显示样式为纵列式时，分多少列显示
'        MaxPerLine ---- 每行显示多少个二级栏目
'返回值：栏目导航的HTML代码
'=================================================
Function GetClass_Navigation(ShowType, Cols, MaxPerLine)
    Dim rsNavigation, rsNavigation2, sqlNavigation, strNavigation, i, j, Class_MenuTitle, strClassUrl
    Dim OpenType_Class
    If Cols <= 0 Then Cols = 1
    If MaxPerLine <= 0 Then MaxPerLine = 3
    sqlNavigation = "select ClassID,ClassName,Depth,ParentID,RootID,LinkUrl,Child,Readme,ClassType,ParentDir,ClassDir,OpenType,ClassPurview from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and ClassType=1 order by RootID,OrderID"
    Set rsNavigation = Conn.Execute(sqlNavigation)
    If rsNavigation.BOF And rsNavigation.EOF Then
        GetClass_Navigation = "没有任何栏目"
    Else
        strNavigation = "<table border='0' cellpadding='0' cellspacing='2'><tr>"
        i = 0
        Do While Not rsNavigation.EOF
            If ShowType = 1 Then
                strNavigation = strNavigation & "<td valign='top' nowrap>"
            Else
                strNavigation = strNavigation & "<td valign='top'"
                If Cols = 2 Then strNavigation = strNavigation & " width='50%'"
                strNavigation = strNavigation & "><table border='0' cellpadding='0' cellspacing='2'><tr><td valign='top'>"
            End If
            If Trim(rsNavigation(7)) <> "" Then
                Class_MenuTitle = Replace(Replace(Replace(Replace(rsNavigation(7), "'", ""), """", ""), Chr(10), ""), Chr(13), "")
            Else
                Class_MenuTitle = ""
            End If
            If rsNavigation(11) = 0 Then
                OpenType_Class = "_self"
            Else
                OpenType_Class = "_blank"
            End If
            If rsNavigation("ClassType") = 2 Then
                strClassUrl = rsNavigation("LinkUrl")
            Else
                strClassUrl = GetClassUrl(rsNavigation("ParentDir"), rsNavigation("ClassDir"), rsNavigation("ClassID"), rsNavigation("ClassPurview"))
            End If
            strNavigation = strNavigation & "【<a class='LinkNavigation' href='" & strClassUrl & "' title='" & Class_MenuTitle & "' target='" & OpenType_Class & "'>" & rsNavigation(1) & "</a>】"
            If ShowType = 1 Then
                strNavigation = strNavigation & "</td><td valign='top'>"
            Else
                strNavigation = strNavigation & "</td></tr><tr><td valign='top'>"
            End If
            
            sqlNavigation = "select ClassID,ClassName,Depth,ParentID,RootID,LinkUrl,Child,Readme,ClassType,ParentDir,ClassDir,OpenType,ClassPurview from PE_Class where ChannelID=" & ChannelID & " and ParentID=" & rsNavigation(0) & " order by OrderID"
            Set rsNavigation2 = Conn.Execute(sqlNavigation)
            j = 0
            Do While Not rsNavigation2.EOF
                If j > 0 Then
                    If j Mod MaxPerLine = 0 Then
                        strNavigation = strNavigation & "<br>"
                    Else
                        If ShowType = 1 Then
                            strNavigation = strNavigation & "&nbsp;&nbsp;&nbsp;"
                        Else
                            strNavigation = strNavigation & " | "
                        End If
                    End If
                End If
                If Trim(rsNavigation2(7)) <> "" Then
                    Class_MenuTitle = Replace(Replace(Replace(Replace(rsNavigation2(7), "'", ""), """", ""), Chr(10), ""), Chr(13), "")
                Else
                    Class_MenuTitle = ""
                End If
                If rsNavigation2(11) = 0 Then
                    OpenType_Class = "_self"
                Else
                    OpenType_Class = "_blank"
                End If
                
                If rsNavigation2("ClassType") = 2 Then
                    strClassUrl = rsNavigation2("LinkUrl")
                Else
                    strClassUrl = GetClassUrl(rsNavigation2("ParentDir"), rsNavigation2("ClassDir"), rsNavigation2("ClassID"), rsNavigation2("ClassPurview"))
                End If
                strNavigation = strNavigation & "<a class='LinkNavigation' href='" & strClassUrl & "' title='" & Class_MenuTitle & "' target='" & OpenType_Class & "'>" & rsNavigation2(1) & "</a>"
                j = j + 1
                rsNavigation2.MoveNext
            Loop
            rsNavigation2.Close
            Set rsNavigation2 = Nothing
            If ShowType = 1 Then
                strNavigation = strNavigation & "</td></tr><tr>"
            Else
                strNavigation = strNavigation & "</td></tr></table>"
                i = i + 1
                If i Mod Cols = 0 Then
                    strNavigation = strNavigation & "</td></tr><tr>"
                Else
                    strNavigation = strNavigation & "</td>"
                End If
            End If
            rsNavigation.MoveNext
        Loop
        strNavigation = strNavigation & "</tr></table>"
    End If
    rsNavigation.Close
    Set rsNavigation = Nothing
    
    GetClass_Navigation = strNavigation
End Function

%>
