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
	End If
	
	
	Call CloseConn()
	
	If bReturn = True Then
		WriteSuccessMsg "����ɹ�!" , "main.asp"
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>δ֪����!</li>"
		End If
	End If
	
	Function Delete()
		Dim lngId
		Dim rsInfo, strSql 

		lngId = ConvertLong(Request("tbId") & "")
		
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From message_t Where regid = " & lngId & ""
		If rsInfo.State = 1 Then rsInfo.Close
		rsInfo.Open strSql, conn, 1, 1
		
		If rsInfo.Eof Then
			strMsg = strMsg & "<br><li>������Ϣ�����ڻ����Ѿ���ɾ����</li>" & vbCrLf
			If rsInfo.State = 1 Then rsInfo.Close
			Set rsInfo = Nothing
			Exit Function
		End IF
		
		If rsInfo.State = 1 Then rsInfo.Close
		Set rsInfo = Nothing

		On Error Resume Next 
		
		strSql = "Delete From message_t Where regid = " & lngId & ""
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'Լ����ͻ
				strMsg = strMsg & "<br><li>��Ϣ�ѱ�ʹ�ã�</li>" & vbCrLf
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
