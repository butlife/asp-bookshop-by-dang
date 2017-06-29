<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%Const Purview_FuncName = "All"%>

<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<%
	Dim strMsg, bReturn, sType
	strMsg = ""
	sType = Trim(Request("type") & "")
	If (sType = "delete") Then
		bReturn = Delete
	ElseIf(sType = "passed") Then
		bReturn = setpass
	End If
	
	
	Call CloseConn()
	
	If bReturn = True Then
		WriteSuccessMsg "处理成功!" , "main.asp"
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>未知错误!</li>"
		End If
	End If

	Function setpass()
		Dim dblId, bLock
		Dim rsAdmin, strSql 

		dblId = ConvertDouble(Request("tbId") & "")
		
		Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select ispassed From manager_t Where IsFixed = false and id = " & dblId & ""
		If rsAdmin.State = 1 Then rsAdmin.Close
		rsAdmin.Open strSql, conn, 1, 1
		
		If (rsAdmin.Bof Or rsAdmin.Eof) Then
			strMsg = strMsg & "<br><li>原因：\n管理员不存在或者已经被删除！\n系统保留管理员禁止删除！</li>" & vbCrLf
			If rsAdmin.State = 1 Then rsAdmin.Close
			Set rsAdmin = Nothing
			Exit Function
		Else
			bLock = ConvertDouble(rsAdmin("ispassed") & "")
		End IF
		
		If rsAdmin.State = 1 Then rsAdmin.Close
		Set rsAdmin = Nothing

		On Error Resume Next 
		
		If (bLock = 1) Then
			strSql = "Update manager_t Set ispassed = 0 Where id = " & dblId & ""
		Else
			strSql = "Update manager_t Set ispassed = 1 Where id = " & dblId & ""
		End If
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>管理员已被使用！</li>" & vbCrLf
			End If
			Err.Clear
			setpass = False
			Exit Function
		End If
		
		setpass = True
	End Function
	
	
	Function Delete()
		Dim lngId
		Dim rsAdmin, strSql 

		lngId = ConvertLong(Request("tbId") & "")
		
		Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From manager_t Where id = " & lngId & ""
		If rsAdmin.State = 1 Then rsAdmin.Close
		rsAdmin.Open strSql, conn, 1, 1
		
		If rsAdmin.Eof Then
			strMsg = strMsg & "<br><li>管理员不存在或者已经被删除！</li>" & vbCrLf
			If rsAdmin.State = 1 Then rsAdmin.Close
			Set rsAdmin = Nothing
			Exit Function
		End IF
		
		If (rsAdmin("IsFixed") = True) Then
			strMsg = strMsg & "<br><li>系统固定账号,不可以删除！</li>" & vbCrLf
			If rsAdmin.State = 1 Then rsAdmin.Close
			Set rsAdmin = Nothing
			Exit Function
		End If
		
		If rsAdmin.State = 1 Then rsAdmin.Close
		Set rsAdmin = Nothing

		On Error Resume Next 
		
		strSql = "Delete From manager_t Where id = " & lngId & ""
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>管理员已被使用！</li>" & vbCrLf
			End If
			Err.Clear
			Delete = False
			Exit Function
		End If
		
		Delete = True
		
	End Function
%>
</body>
</html>
