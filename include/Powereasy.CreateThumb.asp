<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

If Watermark_Position = "" Then Watermark_Position = "1"
If PhotoQuality < 50 Then PhotoQuality = 90
If PhotoQuality > 100 Then PhotoQuality = 90
If Thumb_BackgroundColor = "" Then Thumb_BackgroundColor = "#CCCCCC"
Watermark_Images_Transparence = Watermark_Images_Transparence / 100
Watermark_Text_FontColor = "&H" & Replace(Right(Watermark_Text_FontColor, 6), "#", "")
Watermark_Images_BackgroundColor = "&H" & Replace(Right(Watermark_Images_BackgroundColor, 6), "#", "")
Thumb_BackgroundColor = "&H" & Replace(Right(Thumb_BackgroundColor, 6), "#", "")

Class CreateThumb

'=================================================
'AddWatermark
'作  用：根据选择的外部组件调用图像处理功能给图片加上水印
'参  数：ImgFileName     ---- 图片路径
'=================================================
Public Function AddWatermark(ImgFileName)
    Dim objFont, FileExt
    Dim iLeft, iTop
    Dim LogoWidth, LogoHeight

    AddWatermark = False

    If PhotoObject <= 0 Then Exit Function

    FileExt = GetPhotoExt(ImgFileName)
    If FileExt <> "jpg" And FileExt <> "jpeg" And FileExt <> "jpe" And FileExt <> "bmp" And FileExt <> "gif" Then Exit Function
    
    On Error Resume Next
    
    Select Case PhotoObject
    Case 1       'AspJpegV1.5

        If IsObjInstalled("Persits.Jpeg") = False Then Exit Function
        
        Dim AspJpeg
        Set AspJpeg = Server.CreateObject("Persits.Jpeg")
        AspJpeg.Open Trim(Server.MapPath(ImgFileName))
        If AspJpeg.OriginalWidth > Watermark_Position_X * 2 Then
            If Watermark_Type = 0 Then
                If Watermark_Text <> "" And Watermark_Text_FontColor <> "" Then
                    LogoWidth = (Watermark_Text_FontSize + 1) * GetStrLen(Watermark_Text) / 2
                    LogoHeight = Watermark_Text_FontSize + 1

                    iLeft = GetPosition_X(AspJpeg.OriginalWidth, LogoWidth, Watermark_Position_X)
                    iTop = GetPosition_Y(AspJpeg.OriginalHeight, LogoHeight, Watermark_Position_Y)

                    AspJpeg.Canvas.Font.COLOR = Watermark_Text_FontColor         ' 文字的颜色
                    AspJpeg.Canvas.Font.Family = Watermark_Text_FontName         ' 文字的字体
                    AspJpeg.Canvas.Font.size = Watermark_Text_FontSize           ' 文字的大小
                    AspJpeg.Canvas.Font.Bold = Watermark_Text_Bold               ' 文字是否粗体
                    AspJpeg.Canvas.Font.Quality = 4                              ' Antialiased
                    AspJpeg.Canvas.PrintText iLeft, iTop, Watermark_Text         ' 加入文字的位置坐标
                    AspJpeg.Canvas.Pen.COLOR = &H0               ' 边框的颜色
                    AspJpeg.Canvas.Pen.Width = 1                 ' 边框的粗细
                    AspJpeg.Canvas.Brush.Solid = False           ' 图片边框内是否填充颜色
                    AspJpeg.Quality = PhotoQuality
                    AspJpeg.save Server.MapPath(ImgFileName)     ' 生成文件
                End If
            Else

                If Not fso.FileExists(Server.MapPath(Watermark_Images_FileName)) Then
                    Exit Function
                End If

                Dim AspJpeg2
                Set AspJpeg2 = Server.CreateObject("Persits.Jpeg")
                AspJpeg2.Open Server.MapPath(Watermark_Images_FileName)  '打开水印图片
                iLeft = GetPosition_X(AspJpeg.OriginalWidth, AspJpeg2.Width, Watermark_Position_X)
                iTop = GetPosition_Y(AspJpeg.OriginalHeight, AspJpeg2.Height, Watermark_Position_Y)
                AspJpeg.DrawImage iLeft, iTop, AspJpeg2, Watermark_Images_Transparence, Watermark_Images_BackgroundColor, 90 '在原图上添加水印图片
                AspJpeg.Quality = PhotoQuality
                AspJpeg.save Server.MapPath(ImgFileName)
                Set AspJpeg2 = Nothing
            End If
        End If
        Set AspJpeg = Nothing
    Case 2

    Case 3

    End Select

    AddWatermark = True
    If Err Then
        Err.Clear
        CreateThumb = False
    End if
End Function

'=================================================
'过程名：CreateThumb
'作  用：根据选择的外部组件调用图像处理功能（缩略图，水印）
'参  数：ImgFileName     ----原始图片路径
'        ThumbFileName  ----创建缩略图保存的路径
'        ImageWidth  ----缩略图宽度
'        ImageHeight ----缩略图高度
'=================================================
Public Function CreateThumb(ImgFileName, ThumbFileName, ImageWidth, ImageHeight)
    Dim FileExt, bl_w, bl_h
    Dim iLeft, iTop

    CreateThumb = False

    If PhotoObject <= 0 Then Exit Function
    If ImageWidth = 0 And ImageHeight = 0 Then
        ImageWidth = Thumb_DefaultWidth
        ImageHeight = Thumb_DefaultHeight
    End If

    FileExt = GetPhotoExt(ImgFileName)

    If FileExt <> "jpg" And FileExt <> "jpeg" And FileExt <> "jpe" And FileExt <> "bmp" And FileExt <> "gif" Then Exit Function
    
    On Error Resume Next
    
    Select Case PhotoObject
    Case 1       'AspJpegV1.5

        If IsObjInstalled("Persits.Jpeg") = False Then Exit Function
        
        Dim AspJpeg, AspJpeg2

        Set AspJpeg = Server.CreateObject("Persits.Jpeg")
        Set AspJpeg2 = Server.CreateObject("Persits.Jpeg")
        AspJpeg.Open Trim(Server.MapPath(ImgFileName))
        AspJpeg2.Open Trim(Server.MapPath(ImgFileName))
        
        bl_w = ImageWidth / AspJpeg.OriginalWidth
        bl_h = ImageHeight / AspJpeg.OriginalHeight
        
        If ImageWidth > 0 Then
            If ImageHeight > 0 Then
                Select Case Thumb_Arithmetic
                Case 0   '常规算法：宽度和高度都大于0时，直接缩小成指定大小，其中一个为0时，按比例缩小
                    If bl_w < 1 Or bl_h < 1 Then
                        AspJpeg.Width = ImageWidth
                        AspJpeg.Height = ImageHeight
                        AspJpeg.Quality = PhotoQuality
                        AspJpeg.save Server.MapPath(ThumbFileName)
                        CreateThumb = True
                    End If
                Case 1    '裁剪法：宽度和高度都大于0时，先按最佳比例缩小再裁剪成指定大小，其中一个为0时，按比例缩小
                    If bl_w < 1 Or bl_h < 1 Then
                        If bl_w < bl_h Then
                            AspJpeg.Height = ImageHeight
                            AspJpeg.Width = Round(AspJpeg.OriginalWidth * bl_h)   '按缩小成大比例者
                        Else
                            AspJpeg.Width = ImageWidth
                            AspJpeg.Height = Round(AspJpeg.OriginalHeight * bl_w)
                        End If
                        AspJpeg.Crop 0, 0, ImageWidth, ImageHeight
                        AspJpeg.Quality = PhotoQuality
                        AspJpeg.save Server.MapPath(ThumbFileName)
                        CreateThumb = True
                    End If
                Case 2  '补充法：在指定大小的背景图上附加上按最佳比例缩小的图片
                    
                    '创建一个指定大小的背景图
                    AspJpeg2.Width = ImageWidth
                    AspJpeg2.Height = ImageHeight
                    AspJpeg2.Canvas.Brush.Solid = True            ' 图片边框内是否填充颜色
                    AspJpeg2.Canvas.Brush.COLOR = Thumb_BackgroundColor  '设定背景颜色
                    AspJpeg2.Canvas.Bar -1, -1, AspJpeg2.Width + 1, AspJpeg2.Height + 1 '填充

                    '按最佳比例缩小图片
                    If bl_w > bl_h Then
                        If bl_h < 1 Then
                            AspJpeg.Height = ImageHeight
                            AspJpeg.Width = Round(AspJpeg.OriginalWidth * bl_h)   '按缩小成小比例者
                        End If
                    Else
                        If bl_w < 1 Then
                            AspJpeg.Width = ImageWidth
                            AspJpeg.Height = Round(AspJpeg.OriginalHeight * bl_w)
                        End If
                    End If

                    '得到缩略图的坐标
                    iLeft = (AspJpeg2.Width - AspJpeg.Width) / 2
                    iTop = (AspJpeg2.Height - AspJpeg.Height) / 2

                    AspJpeg2.DrawImage iLeft, iTop, AspJpeg   '将缩略图附加到背景上
                    AspJpeg2.Quality = PhotoQuality
                    AspJpeg2.save Server.MapPath(ThumbFileName)
                    CreateThumb = True
                End Select

            Else
                If bl_w < 1 Then
                    AspJpeg.Width = ImageWidth
                    AspJpeg.Height = Round(AspJpeg.OriginalHeight * bl_w)
                    AspJpeg.Quality = PhotoQuality
                    AspJpeg.save Server.MapPath(ThumbFileName)
                    CreateThumb = True
                End If
            End If

        Else
            If ImageHeight > 0 And bl_h < 1 Then
                AspJpeg.Height = ImageHeight
                AspJpeg.Width = Round(AspJpeg.OriginalWidth * bl_h)
                AspJpeg.Quality = PhotoQuality
                AspJpeg.save Server.MapPath(ThumbFileName)
                CreateThumb = True
            Else
                '宽度和高度都为0时，不做任何处理
            End If
        End If
        Set AspJpeg = Nothing
        Set AspJpeg2 = Nothing

    Case 2

    Case 3

    End Select

    If Err Then
        Err.Clear
        CreateThumb = False
    End if
End Function

Private Function GetPosition_X(xImage_W, xLogo_W, SpaceVal)
    Select Case Watermark_Position
    Case 0 '左上
        GetPosition_X = SpaceVal
    Case 1 '左下
        GetPosition_X = SpaceVal
    Case 2 '居中
        GetPosition_X = (xImage_W - xLogo_W) / 2
    Case 3 '右上
        GetPosition_X = xImage_W - xLogo_W - SpaceVal
    Case 4 '右下
        GetPosition_X = xImage_W - xLogo_W - SpaceVal
    Case Else '不显示
        GetPosition_X = 0
End Select

End Function

Private Function GetPosition_Y(yImage_H, yLogo_H, SpaceVal)
    Select Case Watermark_Position
    Case 0 '左上
        GetPosition_Y = SpaceVal
    Case 1 '左下
        GetPosition_Y = yImage_H - yLogo_H - SpaceVal
    Case 2 '居中
        GetPosition_Y = (yImage_H - yLogo_H) / 2
    Case 3 '右上
        GetPosition_Y = SpaceVal
    Case 4 '右下
        GetPosition_Y = yImage_H - yLogo_H - SpaceVal
    Case Else '不显示
        GetPosition_Y = 0
    End Select

End Function

'取得文件的后缀名
Private Function GetPhotoExt(FullPath)
    Dim strFileExt

    If FullPath <> "" Then
        strFileExt = ReplaceBadChar(Trim(LCase(Mid(FullPath, InStrRev(FullPath, ".") + 1))))

        If Len(strFileExt) > 10 Then
            GetPhotoExt = Left(strFileExt, 3)
        Else
            GetPhotoExt = strFileExt
        End If

    Else
        GetPhotoExt = ""
    End If

End Function

End Class
%>
