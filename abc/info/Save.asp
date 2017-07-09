<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%Const Purview_FuncName = "All"%>

<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!-- #include file="../../common/message.asp"-->
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
		Dim lngsortid, lnginfoid, strtitle, dtmakedate, ispassed, istop, strauthor, sType, sContent, hit, adminid, i, iorder, strKeyWords, strpicurl, strremark
		Dim rsInfo, strSql
		dim lnginfosort_CBS_ID, lnginfosort_FM_ID, lnginfosort_NL_ID, lnginfosort_XL_ID, lnginfosort_ZT_ID, lngiCount

		lngiCount = ConvertLong(Request("iCount") & "")
		lnginfosort_CBS_ID = ConvertLong(Request("infosort_CBS_ID") & "")
		lnginfosort_FM_ID = ConvertLong(Request("infosort_FM_ID") & "")
		lnginfosort_NL_ID = ConvertLong(Request("infosort_NL_ID") & "")
		lnginfosort_XL_ID = ConvertLong(Request("infosort_XL_ID") & "")
		lnginfosort_ZT_ID = ConvertLong(Request("infosort_ZT_ID") & "")
		lnginfoid = ConvertLong(Request("infoid") & "")
		strtitle = Trim(Request("title") & "")
		ispassed = ConvertLong(Request("ispassed") & "")
		istop = ConvertLong(Request("istop") & "")
		sType = trim(request("savetype") & "")
		dtmakedate = trim(request("makedate") & "")
		strKeyWords = trim(request("KeyWords") & "")
		strauthor = trim(request("author") & "")
		hit = ConvertLong(request("hit") & "")
		iorder = ConvertLong(request("iorder") & "")
		strpicurl = trim(request("picurl") & "")
		adminid = ConvertLong(request.cookies(gstrSessionPrefix & "adminid") & "")
		'strremark = DBC2SBC(Request("remark"),0)
		strremark = Request("remark")
		'======开始：eWebEditor编辑区取值=============
		sContent = Request("s_News")
		'sContent = replaceCode(Request("s_News"))
		'=============================================		
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		'On Error Resume Next
		conn.BeginTrans	'开始
		'////////////////////////////////
'		strSql = "Select * From infosort_t Where sortid = " & lngsortid
'		If rsInfo.State = 1 Then rsInfo.Close
'		rsInfo.Open strSql, conn, 1, 1
'		
'		If (rsInfo.Eof Or rsInfo.Bof) Then
'			strMsg = "该新闻类别不存在或己被删除!"
'			infosave = False
'			If rsInfo.State = 1 Then rsInfo.Close
'			Set rsInfo = Nothing
'			Exit Function
'		End If
		'////////////////////////////////

		If (sType = "add") Then
			strSql = "Select * From info_t Where 1 = 2"
			If rsInfo.State = 1 Then rsInfo.Close
			rsInfo.Open strSql, conn, 1, 3
			rsInfo.AddNew
		ElseIf (sType = "modify") then
			strSql = "Select * From info_t Where InfoID = " & lnginfoid
			If rsInfo.State = 1 Then rsInfo.Close
			rsInfo.Open strSql, conn, 1, 3
		End If
		If Not(rsInfo.Bof Or rsInfo.Eof) Then
			rsInfo("iCount") = lngiCount
			rsInfo("infosort_FM_ID") = lnginfosort_FM_ID
			rsInfo("infosort_CBS_ID") = lnginfosort_CBS_ID
			rsInfo("infosort_NL_ID") = lnginfosort_NL_ID
			rsInfo("infosort_XL_ID") = lnginfosort_XL_ID
			rsInfo("infosort_ZT_ID") = lnginfosort_ZT_ID
			rsInfo("adminid") = adminid
			rsInfo("title") = strtitle
			rsInfo("istop") = istop
			rsInfo("ispassed") = ispassed
			rsInfo("makedate") = dtmakedate
			rsInfo("author") = strauthor
			rsInfo("KeyWords") = strKeyWords
			rsInfo("content") = sContent
			rsInfo("remark") = strremark
			rsInfo("hit") = hit
			rsInfo("picurl") = strpicurl
			rsInfo("iorder") = iorder
			rsInfo.Update
		Else
			strMsg = "找不到该图书信息!"
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
