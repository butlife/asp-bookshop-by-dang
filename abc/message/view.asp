<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<style>
	td {
		padding:5px;
	}
</style>
</head>
<body style="padding:10px;">
<%
	Dim rsreg, strSql, lngregid
	Dim strregname, strsex, strage, strtelephone, strQQ, strmail, straddress, strremark, strIP, strMakedate, strtype, strMobtel
	
	lngregid = int(Request("ID") & "")
	
	Set rsreg = Server.CreateObject("ADODB.RecordSet")
	strSql = "Select * From Message_T Where regid = " & lngregid
	rsreg.Open strSql, conn, 2, 3
	
	If (rsreg.Bof Or rsreg.Eof) Then 
		response.write "<script>alert('该留言信息不存在或己被删除!'); window.close();</script>"
		response.end
	Else
		strregname = trim(rsreg("regname") & "")
		strTelephone = trim(rsreg("telephone") & "")
		strQQ = trim(rsreg("oicq") & "")
		strmail = trim(rsreg("mail") & "")
		straddress = trim(rsreg("address") & "")
		strremark = trim(rsreg("remark") & "")
		strIP = trim(rsreg("IP") & "")
		strMakedate = trim(rsreg("Makedate") & "")
		strtype = trim(rsreg("stype") & "")
		strMobtel = trim(rsreg("Mobtel") & "")
	end if
%>
<div>
	<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
	  <tr>
		<td width="60" height="25"><strong>姓 名:</strong></td>
		<td><%= TdString(strregname)%></td>
	  </tr>
	  <tr>
		<td width="60" height="25"><strong>联系电话:</strong></td>
		<td><%= TdString(strTelephone)%></td>
	  </tr>
	  <tr>
		<td width="60" height="25"><strong>留言时间:</strong></td>
		<td><%= TdString(strMakedate)%></td>
	  </tr>
	  <tr>
		<td width="60" height="25"><strong>IP地址:</strong></td>
		<td><%= TdString(strIP)%></td>
	  </tr>
<!--	  <tr>
		<td height="25"><strong>移动电话:</strong></td>
		<td><%'= strMobtel%></td>
	  </tr>
	  <tr>
		<td height="25"><strong>E-Mail:</strong></td>
		<td><%'= strmail%></td>
	  </tr>
	  <tr>
		<td height="25"><strong>住 址:</strong></td>
		<td><%'= straddress%></td>
	  </tr>
-->	  <tr>
		<td width="60" height="25"><strong>内容:</strong></td>
		<td><%= strremark%></td>
	  </tr>
  </table>
</div>
</body>
</html>
