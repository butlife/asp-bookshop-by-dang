<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>

<!-- #include file="../common/conn.asp"-->
<!-- #include file="../common/function.asp"-->
<%
	Dim sSessionName, slngId
'	sSessionName = trim(Session(gstrSessionPrefix & "adminname") & "")
'	slngId = trim(Session(gstrSessionPrefix & "adminId") & "")
	sSessionName = trim(request.cookies(gstrSessionPrefix & "adminname") & "")
	slngId = trim(request.cookies(gstrSessionPrefix & "adminId") & "")
	
	If (sSessionName = "" Or slngId = "") Then
		Response.Redirect("login.asp")
	Else
		Response.Redirect("frame.asp")
	End If
	
%>