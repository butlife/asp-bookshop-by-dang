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
		WriteSuccessMsg "����ɹ�!", "main.asp"
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>δ֪����!</li>"
		End If
	End If
	
	Function adminsave()
		Dim lngSortId, strSortName, strRemark, sType, iorder
		Dim rsSort, strSql 

		lngSortId = ConvertLong(Request("SortId") & "")
		iorder = ConvertLong(Request("iorder") & "")
		strSortName = Trim(Request("SortName") & "")
		strRemark = Trim(Request("Remark") & "")
		sType = trim(request("savetype") & "")

		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		
		On Error Resume Next
		conn.BeginTrans	'��ʼ
		strSql = "Select * From infosort_FM Where SortName = '" & strSortName & "' and sortid <> " & lngSortId
		If rsSort.State = 1 Then rsSort.Close
		rsSort.Open strSql, conn, 1, 1
		
		If Not (rsSort.Eof Or rsSort.Bof) Then
			strMsg = "��Ϣ��𼺴���!"
			adminsave = False
			If rsSort.State = 1 Then rsSort.Close
			Set rsSort = Nothing
			Exit Function
		End If
		If (sType = "add") Then
			strSql = "Select * From infosort_FM Where 1 = 2"
			If rsSort.State = 1 Then rsSort.Close
			rsSort.Open strSql, conn, 2, 3
			rsSort.AddNew
		ElseIf (sType = "modify") then
			strSql = "Select * From infosort_FM Where sortid = " & lngSortId
			If rsSort.State = 1 Then rsSort.Close
			rsSort.Open strSql, conn, 2, 3
		End If
		If Not(rsSort.Bof Or rsSort.Eof) Then
			rsSort("SortName") = strSortName
			rsSort("iorder") = iorder
			rsSort("remark") = strRemark
			rsSort.Update
		Else
			strMsg = "�Ҳ�������Ϣ���!"
			Err.Clear
			conn.RollBackTrans	'���ִ���ع�����
			adminsave = False
			If rsSort.State = 1 Then rsSort.Close
			Set rsSort = Nothing
			Exit Function
		End If
		If Err Then
			Err.Clear
			conn.RollBackTrans	'���ִ���ع�����
			adminsave = False
			If rsSort.State = 1 Then rsSort.Close
			Set rsSort = Nothing
			Exit Function
		End If
	
		conn.CommitTrans	'û�д���,�ύ����
		
		adminsave = True
		If rsSort.State = 1 Then rsSort.Close
		Set rsSort = Nothing
		
	End Function
%>
</body>
</html>
