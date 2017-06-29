<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%Const Purview_FuncName = "All"%>
<%Response.Charset = "GB2312"%>
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
		WriteSuccessMsg "保存成功!", "main.asp"
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>未知错误!</li>"
		End If
	End If
	
	Function adminsave()
		Dim lngUserId, struserName, struserAcc, struserAdd, struserSex, struserpwd, strusertel, strexpdate_s, strexpdate_e, strRemark, sType, lnguseCounts, lngmaxuseCounts, strhobby, lngispassed
		Dim rsUser, strSql 

		lngUserId = ConvertLong(Request("UserId") & "")
		lngispassed = ConvertLong(Request("ispassed") & "")
		struserName = Trim(Request("userName") & "")
		struserAcc = Trim(Request("userAcc") & "")
		struserAdd = Trim(Request("userAdd") & "")
		struserSex = Trim(Request("userSex") & "")
		struserpwd = Trim(Request("userpwd") & "")
		strusertel = Trim(Request("usertel") & "")
		strexpdate_s = Format_Time(Request("expdate_s"),2)
		strexpdate_e = Format_Time(Request("expdate_e"),2)
		lnguseCounts = ConvertLong(Request("useCounts") & "")
		lngmaxuseCounts = ConvertLong(Request("maxuseCounts") & "")
		strhobby = Trim(Request("hobby") & "")
		strRemark = Trim(Request("Remark") & "")
		sType = trim(request("savetype") & "")

		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		
		'On Error Resume Next
		conn.BeginTrans	'开始
		If (sType = "add") Then
			strSql = "Select userAcc From user_t Where userAcc = '" & struserAcc & "'"
			If rsUser.State = 1 Then rsUser.Close
			rsUser.Open strSql, conn, 1, 1
			
			If Not(rsUser.Eof Or rsUser.Bof) Then
				strMsg = "帐号己存在!" & strSql
				adminsave = False
				If rsUser.State = 1 Then rsUser.Close
				Set rsUser = Nothing
				Exit Function
			End If
			
			strSql = "Select * From user_t Where 1 = 2"
			If rsUser.State = 1 Then rsUser.Close
			rsUser.Open strSql, conn, 2, 3
			rsUser.AddNew
		ElseIf (sType = "modify") then
			strSql = "Select * From user_t Where userid = " & lngUserId
			If rsUser.State = 1 Then rsUser.Close
			rsUser.Open strSql, conn, 2, 3
		End If
		If Not(rsUser.Bof Or rsUser.Eof) Then
			rsUser("userName") = struserName
			rsUser("userAcc") = struserAcc
			rsUser("userAdd") = struserAdd
			rsUser("userSex") = struserSex
			rsUser("userpwd") = struserpwd
			rsUser("usertel") = strusertel
			rsUser("expdate_s") = strexpdate_s
			rsUser("expdate_e") = strexpdate_e
			rsUser("useCounts") = lnguseCounts
			rsUser("maxuseCounts") = lngmaxuseCounts
			rsUser("maxuseCountsTemp") = lngmaxuseCounts
			rsUser("hobby") = strhobby
			rsUser("ispassed") = lngispassed
			rsUser("remark") = strRemark
			rsUser.Update
		Else
			strMsg = "找不到该会员!"
			Err.Clear
			conn.RollBackTrans	'出现错误回滚操作
			adminsave = False
			If rsUser.State = 1 Then rsUser.Close
			Set rsUser = Nothing
			Exit Function
		End If
		If Err Then
			Err.Clear
			conn.RollBackTrans	'出现错误回滚操作
			adminsave = False
			If rsUser.State = 1 Then rsUser.Close
			Set rsUser = Nothing
			Exit Function
		End If
	
		conn.CommitTrans	'没有错误,提交数据
		
		adminsave = True
		If rsUser.State = 1 Then rsUser.Close
		Set rsUser = Nothing
		
	End Function
%>
</body>
</html>
