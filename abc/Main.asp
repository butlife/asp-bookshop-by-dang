<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!-- #include file="../common/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>您的服务器信息</title>
<style type="text/css">
	* { font-size:12px;}
	body { margin:5px auto; padding:0;}
	td { text-indent:10px; height:23px; line-height:23px;}
</style>
</head>
<body>
<table width="96%" border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" align="center">
  <tr>
    <td width="50%" height="20">服务器域名：<font> <%=Request.ServerVariables("server_name")%> / <%=Request.ServerVariables("Http_HOST")%></font></td>
    <td width="50%">脚本解释引擎：<font class="t4"> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></font></td>
  </tr>
  <tr>
    <td width="50%" height="20">服务器软件的名称：<font class="t4"> <%=Request.ServerVariables("SERVER_SOFTWARE")%></font></td>
    <td width="50%">浏览器版本：<font class="t4"> <%=Request.ServerVariables("Http_User_Agent")%></font></td>
  </tr>
</table>
<br>
<%
	Call Discreteness
%>
<br>
<%
	'组件
	Sub Discreteness
%>
<table width="96%" border="1" cellpadding="0" cellspacing="0" style="border-collapse:collapse;" align="center">
  <tr>
    <td width="50%" height="22">组件名称</td>
    <td width="50%" height="22">支持及版本</td>
  </tr>
  <%
	Dim theInstalledObjects(18)
	theInstalledObjects(0) = "MSWC.AdRotator"
	theInstalledObjects(1) = "MSWC.BrowserType"
	theInstalledObjects(2) = "MSWC.NextLink"
	theInstalledObjects(3) = "MSWC.Tools"
	theInstalledObjects(4) = "MSWC.Status"
	theInstalledObjects(5) = "MSWC.Counters"
	theInstalledObjects(6) = "MSWC.PermissionChecker"
	theInstalledObjects(7) = "ADODB.Stream"
	theInstalledObjects(8) = "Adodb.connection"
	theInstalledObjects(9) = "Scripting.FileSystemObject"
	theInstalledObjects(10) = "SoftArtisans.FileUp"
	theInstalledObjects(11) = "SoftArtisans.FileManager"
	theInstalledObjects(12) = "JMail.Message"
	theInstalledObjects(13) = "CDONTS.NewMail"
	theInstalledObjects(14) = "Persits.MailSender"
	theInstalledObjects(15) = "LyfUpload.UploadFile"
	theInstalledObjects(16) = "Persits.Upload.1"
	theInstalledObjects(17) = "W3.upload"
	theInstalledObjects(18) = "Adodb.recordset"
	
	Dim i
	For i= 0 to 18
		Response.Write "<TR><TD>" & theInstalledObjects(i)
		Select Case i
		Case 8, 18
			Response.Write "(ACCESS 数据库)"
		Case 9
			Response.Write "(FSO 文本文件读写)"
		Case 10
			Response.Write "(SA-FileUp 文件上传)"
		Case 11
			Response.Write "(SA-FM 文件管理)"
		Case 12
			Response.Write "(JMail 邮件发送)"
		Case 13
			Response.Write "(WIN虚拟SMTP 发信)"
		Case 14
			Response.Write "(ASPEmail 邮件发送)"
		Case 15
			Response.Write "(LyfUpload 文件上传)"
		Case 16
			Response.Write "(ASPUpload 文件上传)"
		Case 17
			Response.Write "(w3 upload 文件上传)"
		End Select
		
		Response.Write "</td><td>"
		
		If Not IsObjInstalled(theInstalledObjects(i)) Then
			Response.Write "<strong style=""color:red;"">×</strong>"
		Else
			Response.Write "<strong style=""color:green;"">√</strong> " & getver(theInstalledObjects(i)) & ""
		End If
		
		Response.Write "</td></TR>" & vbCrLf
	Next
%>
</table>
<%
End Sub

''''''''''''''''''''''''''''''
	Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		
		Set xTestObj = Nothing
		Err = 0
	End Function
''''''''''''''''''''''''''''''
	Function getver(Classstr)
		On Error Resume Next
		getver=""
		Err = 0
		
		Dim xTestObj
		Set xTestObj = Server.CreateObject(Classstr)
		If 0 = Err Then getver=xtestobj.version
		
		Set xTestObj = Nothing
		Err = 0
	End Function
%>
</body>
</html>
