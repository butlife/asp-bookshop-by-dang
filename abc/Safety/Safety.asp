<!-- #include file = "Purview.asp" -->
<%
	'��ȫ�ж�
	'��һ��������Ա�Ƿ��½
'	If Session(gstrSessionPrefix & "AdminName") = "" Then
	If request.cookies(gstrSessionPrefix & "AdminName") = "" Then
		Call CloseConn
		Response.Write "<script language=""javascript"" type=""text/javascript"">alert(""���糬ʱ������û�е�½��"");parent.location.href=""" & gstrAdminPanelUrl & "Login.asp"";/*NONELOGIN*/</script>"
		Response.End
	End If

	'�ڶ������û��Ե�ǰ�����Ƿ���Ȩ��
%>