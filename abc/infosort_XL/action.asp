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
		Dim rssort, strSql, rsBook, strBookSql

		lngId = ConvertLong(Request("tbId") & "")
		
		Set rsBook = Server.CreateObject("ADODB.RecordSet")
		strBookSql = "Select * From info_t Where infosort_XL_ID = " & lngId & ""
		If rsBook.State = 1 Then rsBook.Close
		rsBook.Open strBookSql, conn, 1, 1
		if not(rsBook.bof or rsBook.eof) then
			strMsg = strMsg & "<br><li>请先删除类别下所有书本，再尝试删除类别！</li>" & vbCrLf
			If rsBook.State = 1 Then rsBook.Close
			Set rsBook = Nothing
			Exit Function
		end if
		
		Set rssort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_XL Where sortid = " & lngId & ""
		If rssort.State = 1 Then rssort.Close
		rssort.Open strSql, conn, 1, 1
		
		If rssort.Eof Then
			strMsg = strMsg & "<br><li>类别不存在或者已经被删除！</li>" & vbCrLf
			If rssort.State = 1 Then rssort.Close
			Set rssort = Nothing
			Exit Function
		End If
		
		If rssort.State = 1 Then rssort.Close
		Set rssort = Nothing

		On Error Resume Next 
		
		strSql = "Delete From infosort_XL Where sortid = " & lngId & ""
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
