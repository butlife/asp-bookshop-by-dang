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
		Dim lngId, strtitle, strRemark, sType, strhttpurl, strpicurl, strmakedate
		Dim rsAds, strSql 

		lngId = ConvertLong(Request("Id") & "")
		strtitle = Trim(Request("title") & "")
		strhttpurl = Trim(Request("httpurl") & "")
		strpicurl = Trim(Request("picurl") & "")
		strRemark = Trim(Request("Remark") & "")
		sType = trim(request("savetype") & "")

		Set rsAds = Server.CreateObject("ADODB.RecordSet")
		
		On Error Resume Next
		conn.BeginTrans	'开始
		strSql = "Select * From link_t Where title = '" & strtitle & "' and Id <> " & lngId
		If rsAds.State = 1 Then rsAds.Close
		rsAds.Open strSql, conn, 1, 1
		
		If Not (rsAds.Eof Or rsAds.Bof) Then
			strMsg = "链接己存在!"
			adminsave = False
			If rsAds.State = 1 Then rsAds.Close
			Set rsAds = Nothing
			Exit Function
		End If
		If (sType = "add") Then
			strSql = "Select * From link_t Where 1 = 2"
			If rsAds.State = 1 Then rsAds.Close
			rsAds.Open strSql, conn, 2, 3
			rsAds.AddNew
		ElseIf (sType = "modify") then
			strSql = "Select * From link_t Where Id = " & lngId
			If rsAds.State = 1 Then rsAds.Close
			rsAds.Open strSql, conn, 2, 3
		End If
		If Not(rsAds.Bof Or rsAds.Eof) Then
			rsAds("title") = strtitle
			rsAds("httpurl") = strhttpurl
			rsAds("picurl") = strpicurl
			rsAds("makedate") = strmakedate
			rsAds("remark") = strRemark
			rsAds("makedate") = now()
			rsAds.Update
		Else
			strMsg = "找不到该链接!"
			Err.Clear
			conn.RollBackTrans	'出现错误回滚操作
			adminsave = False
			If rsAds.State = 1 Then rsAds.Close
			Set rsAds = Nothing
			Exit Function
		End If
		If Err Then
			Err.Clear
			conn.RollBackTrans	'出现错误回滚操作
			adminsave = False
			If rsAds.State = 1 Then rsAds.Close
			Set rsAds = Nothing
			Exit Function
		End If
	
		conn.CommitTrans	'没有错误,提交数据
		
		adminsave = True
		If rsAds.State = 1 Then rsAds.Close
		Set rsAds = Nothing
		
	End Function
%>
</body>
</html>
