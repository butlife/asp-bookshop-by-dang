<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!-- #include file="../Common/Conn.asp" -->
<%
	Call CloseConn()
	
	Session(gstrSessionPrefix & "AdminId") = ""
	Session(gstrSessionPrefix & "AdminName") = ""
	response.cookies(gstrSessionPrefix & "AdminID") = ""
	response.cookies(gstrSessionPrefix & "AdminName") = ""
	
	Response.Redirect "Login.asp"
%>
