<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>

<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<%
	Dim strMsg, bReturn
	strMsg = ""
	
	bReturn = Delete 
	
	Call CloseConn
	
	If bReturn = True Then
		WriteSuccessMsg "����ɾ���ɹ�!", "main.asp"
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>δ֪����!</li>"
		End If
	End If

	Function Delete()
		Dim arynfoId, strSql
		arynfoId = Trim(Request("CHK") & "")
		Delete = False
		
		On Error Resume Next
		conn.BeginTrans	'��ʼ
		
		strSql = "Delete From shop_t Where InfoId In (" & arynfoId & ")"
		conn.Execute strSql
		
		strSql = "Delete From Fav_t Where InfoId In (" & arynfoId & ")"
		conn.Execute strSql
		
		strSql = "Delete From info_t Where InfoId In (" & arynfoId & ")"
		conn.Execute strSql
		
		If Err Then
			Err.Clear
			conn.RollBackTrans	'���ִ���ع�����
			Exit Function
		End If

		conn.CommitTrans	'û�д���,�ύ����
		
		Delete = True
	End Function
%>
