<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%Const Purview_FuncName = "All"%>

<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!-- #include file="../../common/message.asp"-->
<!-- #include file="../../common/MD5.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<%
	Dim  bReturn, strMsg
	
	strMsg = ""
	
	bReturn = adminsave
	
	Call CloseConn()
	
	If bReturn = True Then
		WriteSuccessMsg "����ɹ�!", "main.asp"
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>δ֪����!</li>"
		End If
	End If
	
	Function adminsave()
		Dim lngAdminId, strAdminName, lngpassed, dtAddTime, stradminpwd, strRemark, sType
		Dim rsAdmin, strSql 

		lngadminId = ConvertLong(Request("adminId") & "")
		strAdminName = Trim(Request("AdminName") & "")
		stradminpwd = Trim(Request("AdminPwd") & "")
		strRemark = Trim(Request("Remark") & "")
		lngpassed = ConvertLong(Request("ckbLock") & "")
		sType = trim(request("savetype") & "")
		dtAddTime = Now()

		Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
		
		If stradminpwd <> Trim(Request("AdminPwd2") & "") Then
			strMsg = "������������벻һ��!"
			adminsave = False
			If rsAdmin.State = 1 Then rsAdmin.Close
			Set rsAdmin = Nothing
			Exit Function
		End If
		
		
		On Error Resume Next
		conn.BeginTrans	'��ʼ
		If (sType = "add") Then
			strSql = "Select * From manager_t Where AdminName = '" & strAdminName & "'"
			If rsAdmin.State = 1 Then rsAdmin.Close
			rsAdmin.Open strSql, conn, 1, 1
			
			If Not (rsAdmin.Eof Or rsAdmin.Bof) Then
				strMsg = "�ʺż�����!"
				adminsave = False
				If rsAdmin.State = 1 Then rsAdmin.Close
				Set rsAdmin = Nothing
				Exit Function
			End If
			
			strSql = "Select * From manager_t Where 1 = 2"
			If rsAdmin.State = 1 Then rsAdmin.Close
			rsAdmin.Open strSql, conn, 2, 3
			rsAdmin.AddNew
			rsAdmin("adddate") = now()
		ElseIf (sType = "modify") then
			strSql = "Select * From manager_t Where id = " & lngadminId
			If rsAdmin.State = 1 Then rsAdmin.Close
			rsAdmin.Open strSql, conn, 2, 3
		End If
		If Not(rsAdmin.Bof Or rsAdmin.Eof) Then
			rsAdmin("AdminName") = strAdminName
			rsAdmin("ispassed") = lngpassed
			rsAdmin("remark") = strRemark
			If (sType = "modify" and stradminpwd <> "") Then
				rsAdmin("Adminpwd") = MD5(stradminpwd)
			End if
			If (sType = "add") Then
				rsAdmin("Adminpwd") = MD5(stradminpwd)
			End if
			rsAdmin.Update
		Else
			strMsg = "�Ҳ����ù���Ա!"
			Err.Clear
			conn.RollBackTrans	'���ִ���ع�����
			adminsave = False
			If rsAdmin.State = 1 Then rsAdmin.Close
			Set rsAdmin = Nothing
			Exit Function
		End If
		If Err Then
			Err.Clear
			conn.RollBackTrans	'���ִ���ع�����
			adminsave = False
			If rsAdmin.State = 1 Then rsAdmin.Close
			Set rsAdmin = Nothing
			Exit Function
		End If
	
		conn.CommitTrans	'û�д���,�ύ����
		
		adminsave = True
		If rsAdmin.State = 1 Then rsAdmin.Close
		Set rsAdmin = Nothing
		
	End Function
%>
</body>
</html>
