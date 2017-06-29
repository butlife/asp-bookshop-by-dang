<!--#include file="../Common/Conn.asp"-->
<html>
<head>
<title><%=gstrSiteName%> - 后台管理</title>
</head>
<frameset rows="60,*,20" frameborder="no" framespacing="0" name="frameMain">
  <frame src="Top.asp" name="top" scrolling="no" noresize="noresize">
  <frameset cols="180,10,*" id="frameLeft" name="frameLeft">
	<frame src="Left.asp" name="leftmenu" scrolling="no" noresize="noresize">
	<frame src="Mid.asp" name="menubar" scrolling="no" noresize="noresize">
	<frame src="Main.asp" name="main" scrolling="auto" noresize="noresize">
  </frameset>
  <frame src="Bottom.asp" name="bottom" scrolling="NO" noresize="noresize">
</frameset>
<noframes></noframes>
</html>
