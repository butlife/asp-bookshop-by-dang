<%Response.Charset = "GB2312"%>
<html>
<!-- #include file="class.asp" -->
<!-- #include file="config.asp" -->
<head>
<title>�ϴ��ļ�</title>
<style type="text/css">
	* {font-size:12px;}
	body {font-size:12px; padding:0; margin:0; line-height:20px;}
</style>
</head>
<body>
�����ĵȴ����ļ������ϴ��С�������
<%
dim frm, frminput
frm = trim(request("frm")&"")
frminput = trim(request("frminput")&"")

  Server.ScriptTimeout = 9999
  set Upload = new DoteyUpload
  Upload.MaxTotalBytes = 5 *1024 *1024	' ���10MB

  Upload.Upload() '�ϴ���ʾ�������浽Ӳ��

  If Request.TotalBytes > 10 *1024 *1024 Then
	Response.Write "��Ҫ�ϴ�����2MB���ļ�"
	Response.End
  End If 

  if Upload.ErrMsg <> "" then 
    Response.Write(Upload.ErrMsg)
    Response.End()
  end if

  if Upload.Files.Count > 0 then
	Items = Upload.Files.Items
  end if

'  Response.Write("�����ϴ� " & Upload.Files.Count & " ���ļ���: " & path & "<hr>")
  for each File in Upload.Files.Items
	dim sssFileName, sRnd, sExt, aExt, isext, k
	sRnd = Int(900 * Rnd) + 100
	sExt = lcase(File.FileExt)
	isext = false
	aExt = Split(gstrAllowExt, "|")
	For kkk = 0 To UBound(aExt)
		If LCase(aExt(kkk)) = right(sExt,3) Then
			isext = True
			Exit For
		End If
	Next
	
	if (isext = false) then
		response.write "<script>; alert('ֻ���ϴ� " & gstrAllowExt & " ��ʽ���ļ�');history.back();</script>"
		response.end
	end if
	sssFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & sRnd & sExt
	
	File.SaveAs(gstrUpLoadPath_big & sssFileName)
	'����Сͼ
	'strsdfasdfasd = CreateSmallPic(gstrUpLoadPath_big, gstrUpLoadPath_small, 300, sssFileName)
  next
  

'response.write sssFileName
%>
�ϴ��ɹ����ļ�����<a href="<%= gstrInstallDir%>uppic/big/<%=sssFileName%>" target="_blank"><%=sssFileName%></a>,<a href="<%= gstrInstallDir%>upload/select.asp?frm=form1&frminput=picurl&m=<%=now()%>">�����ϴ�</a>��
<script language="javascript">
	window.parent.document.<%=frm%>.<%=frminput%>.value = "<%=sssFileName%>";
</script>
</body>
</html>
<%  
  
	Function CreateSmallPic(bigpicpath,smallpicpath,iWidth, filename)

		bigpicpath = bigpicpath&filename
		smallpicpath = smallpicpath&filename
		
		if bigpicpath<>"" then
			dim strAllOldPicPath
			strAllOldPicPath = Server.MapPath(bigpicpath)
			randomize '���������
			
			set fso = createobject("scripting.filesystemobject")
			Set Jpeg = Server.CreateObject("Persits.Jpeg") '����ʵ��

			if fso.FileExists(strAllOldPicPath) then	
				Jpeg.Open strAllOldPicPath '��ͼƬ
				Jpeg.Width = iWidth
				Jpeg.Height = iWidth*Jpeg.OriginalHeight/Jpeg.OriginalWidth		'����������
				Jpeg.Save Server.MapPath(smallpicpath)
				Jpeg.Close:Set Jpeg = Nothing
				CreateSmallPic = smallpicpath & filename
			else
				exit function	
			end if	
		end if
	End Function
%>