<%
	call safeChk()
	
	function safeChk()
		If trim(Session("userid")&"") = "0" or Session("useracc") = "" Then
			Call CloseConn
			Response.Write "<script language=""javascript"" type=""text/javascript"">alert(""网络超时或您还没有登陆！"");parent.location.href=""" & gstrInstallDir & "index.asp"";/*NONELOGIN*/</script>"
			Response.End
		End If
	end function
%>