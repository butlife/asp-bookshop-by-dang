<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
</head>
<body style="padding:10px; text-align:center;">
<%
	Dim rsMess, strSql, lngregid
	Dim strregname, stype, strrecontent
	
	lngregid = int(Request("messid") & "")
	stype = trim(Request("stype") & "")
	Set rsMess = Server.CreateObject("ADODB.RecordSet")
	strSql = "Select * From Message_T Where regid = " & lngregid
	rsMess.Open strSql, conn, 2, 3
	
	If (rsMess.Bof Or rsMess.Eof) Then 
		response.write "<script>alert('该留言信息不存在或己被删除!'); window.close();</script>"
		response.end
	Else
		strrecontent = rsmess("recontent")
		if (stype = "save") then
			if (strrecontent = "") then
				response.write "<script>alert('没有回复内容~');</script>"
			else
				strrecontent = trim(Request("recontent") & "")
				rsmess("recontent") = strrecontent
				rsmess.update
				response.write "<script>alert('回复成功 '); window.close();</script>"
			end if
		end if
	end if
%>
<form action="msg.asp?stype=save" method="post" name="msgfrm" id="msgfrm" target="msgframe" onSubmit="if (document.getElementById('recontent').value == '') {alert('没有回复内容~'); return false;}">
  <input type="hidden" id="messid" name="messid" value="<%=lngregid%>" />
  <textarea name="recontent" id="recontent" cols="65" rows="16"><%=strrecontent%></textarea>
  <input type="submit" value="确定回复" />
</form>
<iframe name="msgframe" style="display:none;"></iframe>
</body>
</html>