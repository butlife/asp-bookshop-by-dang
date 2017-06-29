
<% 
'============================================================================= 
'动网论坛加法验证码（ASPJpeg版） 
'作者：cuixiping 
'博客(CSDN)：http://blog.csdn.net/cuixiping/ 
'博客(博客园)：http://www.cnblogs.com/cuixiping/ 
'日期：2008年11月 
'生成100x20的jpeg图片验证码，内容类似“25+64等于？” 
'需要网站空间支持ASPJpeg组件(Persits.Jpeg)。 
'如果服务器没有“楷体_GB2312”字体，请修改为其他字体并调整字符位置。 
'使用方法：将该代码替换动网论坛的Dv_GetCode.asp文件中的内容，文件首尾不要留空行。 
'============================================================================= 

Const FontColor = &HFFFFFF ' 字体颜色 
Const BgColor = &H222222 ' 背景颜色 
Call CreatValidCode("VerificationCode") 

Sub CreatValidCode(PSN)
Dim x, Jpeg 
Randomize 
x = Array(1+Int(Rnd()*9), Int(Rnd()*10), 1+Int(Rnd()*9), Int(Rnd()*10), 0, 0, "+") 
'-----1/2 位数字-------
x(4) = x(0)
x(5) = x(2)

'x(4) = x(0)*10 + x(1) 
'x(5) = x(2)*10 + x(3) 
'-----1/2 位数字-------

Session(PSN) = CStr(x(4) + x(5)) 
Set Jpeg = Server.CreateObject("Persits.Jpeg")
'图片维度设置
'Jpeg.New 85,20,BgColor 
Jpeg.New 85,20,BgColor 
Jpeg.Quality=100 
With Jpeg.Canvas 
.Font.Bold = True 
.Font.Size = 16 
.Font.Rotation = 0
'字体设置
.Font.Family = "Tahoma" 
.Font.Color = FontColor 
'========1/2 位数字======
.PrintText 3, 2, CStr(x(0)) 
.Font.Rotation = -15 
'.PrintText 36, 2, "=" 
.PrintText 16, -2, "加" 
.Font.Rotation = 0 
'.PrintText 15, 2, x(6) 
.PrintText 30, 2, CStr(x(2)) 
'是否有角度
'.PrintText 36, 2, "=" 
.Font.Rotation = 15 
.PrintText 36, 3, "等" 
.Font.Rotation = -15 
.PrintText 58, -2, "于" 
.Font.Rotation = 9 
.PrintText 69, 3, "？" 
'.PrintText 55, 2, "？" 

'.PrintText 4, 2, CStr(x(0)) 
'.PrintText 14, 2, CStr(x(1)) 
'.PrintText 26, 2, x(6) 
'.PrintText 38, 2, CStr(x(2)) 
'.PrintText 48, 2, CStr(x(3)) 
'.Font.Rotation = 15 
'.PrintText 55, 2, "=" 
''.PrintText 55, 3, "等" 
''.PrintText 70, 3, "于" 
''.PrintText 85, 3, "？" 
'.PrintText 70, 2, "？" 
'========1/2 位数字=======
End With 
'禁止缓存 
Response.ContentType = "image/jpeg" 
Response.Expires = -9999 
Response.AddHeader "pragma", "no-cache" 
Response.AddHeader "cache-ctrol", "no-cache" 
Response.AddHeader "Content-Disposition","inline; filename=www_ay2s_com.jpg" 
Jpeg.SendBinary 
Jpeg.Close 
Set Jpeg = Nothing 
End Sub 
%>
