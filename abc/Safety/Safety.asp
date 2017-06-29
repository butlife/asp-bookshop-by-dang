<!-- #include file = "Purview.asp" -->
<%
	'安全判断
	'第一步：管理员是否登陆
'	If Session(gstrSessionPrefix & "AdminName") = "" Then
	If request.cookies(gstrSessionPrefix & "AdminName") = "" Then
		Call CloseConn
		Response.Write "<script language=""javascript"" type=""text/javascript"">alert(""网络超时或您还没有登陆！"");parent.location.href=""" & gstrAdminPanelUrl & "Login.asp"";/*NONELOGIN*/</script>"
		Response.End
	End If

	'第二步：用户对当前操作是否有权限
%>