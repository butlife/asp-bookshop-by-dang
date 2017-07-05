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
	If (sType = "delete") Then
		bReturn = Delete
	ElseIf(sType = "passed") Then
		bReturn = setpass
	ElseIf(sType = "top") Then
		bReturn = settop
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

	Function settop()
		Dim dblId, bLock
		Dim rsInfo, strSql 

		dblId = ConvertDouble(Request("tbId") & "")
		
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select istop From info_t Where infoid = " & dblId & ""
		If rsInfo.State = 1 Then rsInfo.Close
		rsInfo.Open strSql, conn, 1, 1
		
		If (rsInfo.Bof Or rsInfo.Eof) Then
			strMsg = strMsg & "<br><li>信息不存在或者已经被删除！</li>" & vbCrLf
			If rsInfo.State = 1 Then rsInfo.Close
			Set rsInfo = Nothing
			Exit Function
		Else
			bLock = ConvertDouble(rsInfo("istop") & "")
		End IF
		
		If rsInfo.State = 1 Then rsInfo.Close
		Set rsInfo = Nothing

		On Error Resume Next 
		
		If (bLock = 1) Then
			strSql = "Update info_t Set istop = 0 Where infoid = " & dblId & ""
		Else
			strSql = "Update info_t Set istop = 1 Where infoid = " & dblId & ""
		End If
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>信息已被使用！</li>" & vbCrLf
			End If
			Err.Clear
			setpass = False
			Exit Function
		End If
		
		settop = True
	End Function

	Function setpass()
		Dim dblId, bLock
		Dim rsInfo, strSql 

		dblId = ConvertDouble(Request("tbId") & "")
		
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select ispassed From info_t Where infoid = " & dblId & ""
		If rsInfo.State = 1 Then rsInfo.Close
		rsInfo.Open strSql, conn, 1, 1
		
		If (rsInfo.Bof Or rsInfo.Eof) Then
			strMsg = strMsg & "<br><li>信息不存在或者已经被删除！</li>" & vbCrLf
			If rsInfo.State = 1 Then rsInfo.Close
			Set rsInfo = Nothing
			Exit Function
		Else
			bLock = ConvertDouble(rsInfo("ispassed") & "")
		End IF
		
		If rsInfo.State = 1 Then rsInfo.Close
		Set rsInfo = Nothing

		On Error Resume Next 
		
		If (bLock = 1) Then
			strSql = "Update info_t Set ispassed = 0 Where infoid = " & dblId & ""
		Else
			strSql = "Update info_t Set ispassed = 1 Where infoid = " & dblId & ""
		End If
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>信息已被使用！</li>" & vbCrLf
			End If
			Err.Clear
			setpass = False
			Exit Function
		End If
		
		setpass = True
	End Function
	
	Function Delete()
		Dim lngId
		Dim rsInfo, strSql 

		lngId = ConvertLong(Request("tbId") & "")
		
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From info_t Where infoid = " & lngId & ""
		If rsInfo.State = 1 Then rsInfo.Close
		rsInfo.Open strSql, conn, 1, 1
		
		If rsInfo.Eof Then
			strMsg = strMsg & "<br><li>信息不存在或者已经被删除！</li>" & vbCrLf
			If rsInfo.State = 1 Then rsInfo.Close
			Set rsInfo = Nothing
			Exit Function
		End IF
		
'		If (rsInfo("ispassed") = 1) Then
'			strMsg = strMsg & "<br><li>信息处于审核状态,不可以删除！</li>" & vbCrLf
'			If rsInfo.State = 1 Then rsInfo.Close
'			Set rsInfo = Nothing
'			Exit Function
'		End If
		
		If rsInfo.State = 1 Then rsInfo.Close
		Set rsInfo = Nothing

		On Error Resume Next 

		strSql = "Delete From shop_t Where infoId = " & lngId
		conn.Execute strSql
		
		strSql = "Delete From Fav_t Where infoId = " & lngId
		conn.Execute strSql

		strSql = "Delete From info_t Where infoid = " & lngId
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>信息已被使用！</li>" & vbCrLf
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
