
<% 
'============================================================================= 
'������̳�ӷ���֤�루ASPJpeg�棩 
'���ߣ�cuixiping 
'����(CSDN)��http://blog.csdn.net/cuixiping/ 
'����(����԰)��http://www.cnblogs.com/cuixiping/ 
'���ڣ�2008��11�� 
'����100x20��jpegͼƬ��֤�룬�������ơ�25+64���ڣ��� 
'��Ҫ��վ�ռ�֧��ASPJpeg���(Persits.Jpeg)�� 
'���������û�С�����_GB2312�����壬���޸�Ϊ�������岢�����ַ�λ�á� 
'ʹ�÷��������ô����滻������̳��Dv_GetCode.asp�ļ��е����ݣ��ļ���β��Ҫ�����С� 
'============================================================================= 

Const FontColor = &HFFFFFF ' ������ɫ 
Const BgColor = &H222222 ' ������ɫ 
Call CreatValidCode("VerificationCode") 

Sub CreatValidCode(PSN)
Dim x, Jpeg 
Randomize 
x = Array(1+Int(Rnd()*9), Int(Rnd()*10), 1+Int(Rnd()*9), Int(Rnd()*10), 0, 0, "+") 
'-----1/2 λ����-------
x(4) = x(0)
x(5) = x(2)

'x(4) = x(0)*10 + x(1) 
'x(5) = x(2)*10 + x(3) 
'-----1/2 λ����-------

Session(PSN) = CStr(x(4) + x(5)) 
Set Jpeg = Server.CreateObject("Persits.Jpeg")
'ͼƬά������
'Jpeg.New 85,20,BgColor 
Jpeg.New 85,20,BgColor 
Jpeg.Quality=100 
With Jpeg.Canvas 
.Font.Bold = True 
.Font.Size = 16 
.Font.Rotation = 0
'��������
.Font.Family = "Tahoma" 
.Font.Color = FontColor 
'========1/2 λ����======
.PrintText 3, 2, CStr(x(0)) 
.Font.Rotation = -15 
'.PrintText 36, 2, "=" 
.PrintText 16, -2, "��" 
.Font.Rotation = 0 
'.PrintText 15, 2, x(6) 
.PrintText 30, 2, CStr(x(2)) 
'�Ƿ��нǶ�
'.PrintText 36, 2, "=" 
.Font.Rotation = 15 
.PrintText 36, 3, "��" 
.Font.Rotation = -15 
.PrintText 58, -2, "��" 
.Font.Rotation = 9 
.PrintText 69, 3, "��" 
'.PrintText 55, 2, "��" 

'.PrintText 4, 2, CStr(x(0)) 
'.PrintText 14, 2, CStr(x(1)) 
'.PrintText 26, 2, x(6) 
'.PrintText 38, 2, CStr(x(2)) 
'.PrintText 48, 2, CStr(x(3)) 
'.Font.Rotation = 15 
'.PrintText 55, 2, "=" 
''.PrintText 55, 3, "��" 
''.PrintText 70, 3, "��" 
''.PrintText 85, 3, "��" 
'.PrintText 70, 2, "��" 
'========1/2 λ����=======
End With 
'��ֹ���� 
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
