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
		WriteSuccessMsg "批量删除成功!", "main.asp"
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>未知错误!</li>"
		End If
	End If

	Function Delete()
		Dim arynfoId, strSql
		arynfoId = Trim(Request("CHK") & "")
		Delete = False
		
		On Error Resume Next
		conn.BeginTrans	'开始
		
		strSql = "Delete From shop_t Where InfoId In (" & arynfoId & ")"
		conn.Execute strSql
		
		strSql = "Delete From Fav_t Where InfoId In (" & arynfoId & ")"
		conn.Execute strSql
		
		strSql = "Delete From info_t Where InfoId In (" & arynfoId & ")"
		conn.Execute strSql
		
		If Err Then
			Err.Clear
			conn.RollBackTrans	'出现错误回滚操作
			Exit Function
		End If

		conn.CommitTrans	'没有错误,提交数据
		
		Delete = True
	End Function
%>
