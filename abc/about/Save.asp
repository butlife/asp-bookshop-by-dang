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
	
	bReturn = infosave
	
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
	
	Function infosave()
		Dim lngInfoId, strTitle, strAdminName, lngAdminId, dtUpdateTime, strContent, sRemark
		Dim rsInfo, strSql, i, sType, strpicurl

		lngInfoId = ConvertLong(Request("id") & "")
		strTitle = Trim(Request("title") & "")
		sRemark = Request("Remark")
		strpicurl = Request("picurl")
		lngAdminId = ConvertLong(request.cookies(gstrSessionPrefix & "AdminId") & "")
		dtUpdateTime = now
		'======开始：eWebEditor编辑区取值=============
		strContent = Request("s_News")
		'=============================================		
		
		sType = trim(request("savetype") & "")
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		
		On Error Resume Next
		conn.BeginTrans	'开始
		If (sType = "add") Then
			strSql = "Select * From about_t Where 1 = 2"
			If rsInfo.State = 1 Then rsInfo.Close
			rsInfo.Open strSql, conn, 2, 3
			rsInfo.AddNew
		ElseIf (sType = "modify") then
			strSql = "Select * From about_t Where id = " & lngInfoId
			If rsInfo.State = 1 Then rsInfo.Close
			rsInfo.Open strSql, conn, 2, 3
		End If
		If Not(rsInfo.Bof Or rsInfo.Eof) Then
			rsInfo("title") = strTitle
			rsInfo("makedate") = dtUpdateTime
			rsInfo("adminid") = lngAdminId
			rsInfo("Remark") = sRemark
			rsInfo("content") = strcontent
			rsInfo("picurl") = strpicurl
			rsInfo.Update
		Else
			strMsg = "找不到该站点信息!"
			Err.Clear
			conn.RollBackTrans	'出现错误回滚操作
			infosave = False
			If rsInfo.State = 1 Then rsInfo.Close
			Set rsInfo = Nothing
			Exit Function
		End If
		If Err Then
			Err.Clear
			conn.RollBackTrans	'出现错误回滚操作
			infosave = False
			If rsInfo.State = 1 Then rsInfo.Close
			Set rsInfo = Nothing
			Exit Function
		End If
	
		conn.CommitTrans	'没有错误,提交数据
		
		infosave = True
		If rsInfo.State = 1 Then rsInfo.Close
		Set rsInfo = Nothing
		
	End Function
%>
</body>
</html>
