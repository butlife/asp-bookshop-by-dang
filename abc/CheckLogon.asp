<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>

<!-- #include file = "../Common/Conn.asp" -->
<!-- #include file = "../Common/Function.asp" -->
<!-- #include file = "../Common/MD5.asp" -->
<!-- #include file = "../Common/Message.asp" -->
<%
	Dim loginFlag, errMsg
	errMsg = ""
	
	loginFlag = ChkLogin(errMsg)
	
	Call CloseConn()

	If loginFlag = False Then
		WriteErrorMsg errMsg
	Else
		Response.Redirect "frame.asp"
	End If

	
	'============================================
	'��������ChkLogin
	'��  �ܣ�����û�(����Ա)��½
	'        errMsg OUTPUT      ������Ϣ
	'����ֵ����
	'============================================
	Function ChkLogin(errMsg)
		Dim strSql, rsAdmin
		Dim lngAdminID, strAdminName, strPassword, lngDeptId, strDeptName, strNickName, lngHeadShipId, strHeadShipName
		Dim rsOnline, vIP, vAgent, vPage, strNowTime

		strAdminName = ReplaceBadChar(Trim(request("adminname") & ""))
		strPassword = ReplaceBadChar(Trim(Request("adminpwd") & ""))

		vIP = Request.ServerVariables("Remote_Addr")
		vAgent = Request.ServerVariables("HTTP_Admin_AGENT")
		vPage = Request.ServerVariables("HTTP_REFERER")
		strNowTime = Now()
		
		ChkLogin = True
		errMsg = ""
		
		If strAdminName = "" Then
			ChkLogin = False
			errMsg = errMsg & "<br><li>�û�������Ϊ�գ�</li>"
		End If
		If strPassword = "" Then
			ChkLogin = False
			errMsg = errMsg & "<br><li>���벻��Ϊ�գ�</li>"
		End If
		
		If ChkLogin = False Then
			Exit Function
		End If
		
		strPassword = Md5(strPassword)

		Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
		On Error Resume Next
		
		strSql = "Select * From manager_t Where id = 1 and adminname = '" & strAdminName & "' And adminpwd = '" & strPassword & "' And ispassed = 1 "
		rsAdmin.Open strSql, conn, 1, 3
		
		If (rsAdmin.Bof And rsAdmin.Eof) Then
			ChkLogin = False
			errMsg = errMsg & "<br><li>�û������������.</li>"
			Exit Function
		Else
			lngAdminID = Trim(rsAdmin("id") & "")
			strAdminName = Trim(rsAdmin("AdminName") & "")
			Session(gstrSessionPrefix & "AdminID") = lngAdminID
			Session(gstrSessionPrefix & "AdminName") = strAdminName
			response.cookies(gstrSessionPrefix & "AdminID") = lngAdminID
			response.cookies(gstrSessionPrefix & "AdminName") = strAdminName
			rsAdmin("sIP") = vIP
			rsAdmin("logindate") = strNowTime
			rsAdmin("logocount") = rsAdmin("logocount") + 1
			rsAdmin.update
		End If
		
		If Err Or ChkLogin = False Then
			Err.Clear
			Session(gstrSessionPrefix & "AdminID") = ""
			Session(gstrSessionPrefix & "AdminName") = ""
			response.cookies(gstrSessionPrefix & "AdminID") = ""
			response.cookies(gstrSessionPrefix & "AdminName") = ""
			rsAdmin.Close
			Set rsAdmin = Nothing
			Exit Function
		End If
		
		rsAdmin.Close
		Set rsAdmin = Nothing
		
	End Function
%>