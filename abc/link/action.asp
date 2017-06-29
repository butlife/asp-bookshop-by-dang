<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%Const Purview_FuncName = "All"%>
<%Response.Charset = "GB2312"%>
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
	
	Function Delete()
		Dim lngId
		Dim rsAds, strSql 

		lngId = ConvertLong(Request("tbId") & "")
		
		Set rsAds = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From link_t Where Id = " & lngId & ""
		If rsAds.State = 1 Then rsAds.Close
		rsAds.Open strSql, conn, 1, 1
		
		If rsAds.Eof Then
			strMsg = strMsg & "<br><li>不存在或者已经被删除！</li>" & vbCrLf
			If rsAds.State = 1 Then rsAds.Close
			Set rsAds = Nothing
			Exit Function
		End If
		
		If rsAds.State = 1 Then rsAds.Close
		Set rsAds = Nothing

		On Error Resume Next 
		
		strSql = "Delete From link_t Where Id = " & lngId & ""
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'约束冲突
				strMsg = strMsg & "<br><li>类别已被使用！</li>" & vbCrLf
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
