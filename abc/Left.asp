<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!-- #include file="../Common/conn.asp"-->
<!-- #include file="../Common/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type content=text/html; charset=gb2312">
<title>���ܲ˵�</title>
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
<div style="font-size:14px; text-align:center;">��ӭ <span style="color:#CA0000;"><%= request.cookies(gstrSessionPrefix & "adminname")%></span> ��½.</div>
<hr>
<div id="#left_logo">
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>Orders/Main.asp" target="main">��������</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>user/Main.asp" target="main">��Ա����</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>info/Main.asp" target="main">ͼ�����</a></div>
  <br>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_NL/Main.asp" target="main">����������</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_ZT/Main.asp" target="main">����������</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_XL/Main.asp" target="main">ϵ��������</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_FM/Main.asp" target="main">��ĸ������</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>infosort_CBS/Main.asp" target="main">������������</a></div>
  <br>
  <!--<div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>about/Main.asp" target="main">վ�����</a></div>-->
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="<%=gstrAdminPanelUrl%>manager/Main.asp" target="main">����Ա����</a></div>
  <div class="logo" onMouseOver="this.className='logo_Mover';" onMouseOut="this.className='logo';"><a href="Logout.asp" target="_parent">�˳�ϵͳ</a></div>
</div>
</body>
</html>
