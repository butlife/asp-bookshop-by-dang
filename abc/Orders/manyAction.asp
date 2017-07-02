<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>

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
	If (sType = "manyDelete") Then
		bReturn = manyDelete
	ElseIf(sType = "manySend") Then
		bReturn = manySend
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

	Function manySend()
		Dim dblId, bLock
		Dim rsShop, strSql 

		dblId = trim(Request("CHK") & "")
		
		Set rsShop = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select shopstate From shop_t Where shopId in (" & dblId & ")"
		If rsShop.State = 1 Then rsShop.Close
		rsShop.Open strSql, conn, 1, 3
		
		If (rsShop.Bof Or rsShop.Eof) Then
			strMsg = strMsg & "<br><li>订单信息不存在或者已经被删除！</li>" & vbCrLf
			If rsShop.State = 1 Then rsShop.Close
			Set rsShop = Nothing
			Exit Function
		End IF
		
		If rsShop.State = 1 Then rsShop.Close
		Set rsShop = Nothing

		On Error Resume Next 
		
		strSql = "Update shop_t Set shopstate = 1 Where shopId in (" & dblId & ")"
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>订单信息已被使用！</li>" & vbCrLf
			End If
			Err.Clear
			setpass = False
			Exit Function
		End If
		
		manySend = True
	End Function
	
	Function manyDelete()
		Dim dblId
		Dim rsShop, strSql 

		dblId = trim(Request("CHK") & "")
		
		Set rsShop = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select shopstate From shop_t Where shopId in (" & dblId & ")"
		If rsShop.State = 1 Then rsShop.Close
		rsShop.Open strSql, conn, 1, 1
		
		If rsShop.Eof Then
			strMsg = strMsg & "<br><li>订单信息不存在或者已经被删除！</li>" & vbCrLf
			If rsShop.State = 1 Then rsShop.Close
			Set rsShop = Nothing
			Exit Function
		End IF
		
		If rsShop.State = 1 Then rsShop.Close
		Set rsShop = Nothing

		On Error Resume Next 
		
		strSql = "Delete From shop_t Where shopId in (" & dblId & ")"
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>信息已被使用！</li>" & vbCrLf
			End If
			Err.Clear
			Delete = False
			Exit Function
		End If
		
		manyDelete = True
		
	End Function
%>
</body>
</html>
