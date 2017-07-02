<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!-- #include file="../Common/conn.asp"-->
<!-- #include file="../Common/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type content=text/html; charset=gb2312">
<title>功能菜单</title>
<style type="text/css">
* {
	font-size:14px;
}

a {
	color:#353535;
	text-decoration: none;
	font-size:14px;
}

a:hover, a:* {
	color:#003399;
	text-decoration: underline;
}

.logo {
	border:1px #666666 solid;
	background:#eeeeee;
	height:30px;
	line-height:30px;
	text-align:center;
	margin:5px auto 5px auto;
}

.logo_Mover {
	margin:5px auto 5px auto;
	border:1px #333333 solid;
	background:#cccccc;
	height:30px;
	line-height:30px;
	text-align:center;
	font-weight:bold;
}
</style>
</head>
<body>
<div style="font-size:14px; text-align:center;">欢迎 <span style="color:#CA0000;"><%= request.cookies(gstrSessionPrefix & "adminname")%></span> 登陆.</div>
<hr>
<div id="#left_logo">
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>Orders/Main.asp" target="main">订单管理</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>user/Main.asp" target="main">会员管理</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>info/Main.asp" target="main">图书管理</a></div>
  <br>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_NL/Main.asp" target="main">年龄类别管理</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_ZT/Main.asp" target="main">主题类别管理</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_XL/Main.asp" target="main">系列类别管理</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_FM/Main.asp" target="main">父母类别管理</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_CBS/Main.asp" target="main">出版社类别管理</a></div>
  <br>
  <!--<div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>about/Main.asp" target="main">站点管理</a></div>-->
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>manager/Main.asp" target="main">管理员管理</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="Logout.asp" target="_parent">退出系统</a></div>
</div>
</body>
</html>
