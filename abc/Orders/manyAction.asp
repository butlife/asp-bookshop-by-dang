<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>

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
	Dim strMsg, bReturn, sType, strMain
	strMsg = ""
	sType = Trim(Request("type") & "")
	strMain = Trim(Request("Main") & "")
	
	If (sType = "manyDelete") Then
		bReturn = manyDelete
	ElseIf(sType = "manySend") Then
		bReturn = manySend
	ElseIf(sType = "manyFinish") Then
		bReturn = manyFinish
	End If
	
	
	Call CloseConn()
	
	If bReturn = True Then
		select case strMain
			case "0"
			WriteSuccessMsg "����ɹ�!" , "main.asp"
			
			case "1"
			WriteSuccessMsg "����ɹ�!" , "main1.asp"
			
			case "2"
			WriteSuccessMsg "����ɹ�!" , "main2.asp"
			
			case "3"
			WriteSuccessMsg "����ɹ�!" , "main3.asp"
			
		
		end select
	Else
		If strMsg <> "" Then
			WriteErrorMsg strMsg
		Else
			WriteErrorMsg "<br><li>δ֪����!</li>"
		End If
	End If

	Function manySend()
		Dim dblId, bLock, thisUserId
		Dim rsShop, strSql, strUserSql
		
		dblId = trim(Request("shopId") & "")
		thisUserId = ConvertLong(Request("userid") & "")
		if dblId = "" then dblId = "0"
		
		Set rsShop = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select shopstate, infoId, userid, sendDate From shop_t Where shopstate = 0 and shopId in (" & dblId & ")"
		If rsShop.State = 1 Then rsShop.Close
		rsShop.Open strSql, conn, 1, 3
		
		If (rsShop.Bof Or rsShop.Eof) Then
			strMsg = strMsg & "<br><li>������Ϣ�����ڻ����Ѿ���ɾ����</li>" & vbCrLf
			If rsShop.State = 1 Then rsShop.Close
			Set rsShop = Nothing
			Exit Function
		else
			do while not(rsshop.bof or rsshop.eof)
				'���¶���״��
				rsshop("shopstate") = 1
				rsshop("sendDate") = now()
				rsshop.update
				rsshop.movenext
			loop
			'���»�Ա�����ܴ���
			strUserSql = "Update user_t Set maxuseCountsTemp = maxuseCountsTemp -1 Where userid = " & thisUserId
			conn.Execute strUserSql
		End IF
		
		If rsShop.State = 1 Then rsShop.Close
		Set rsShop = Nothing

		On Error Resume Next 

		If Err Then
			If Err.Number = -2147217900 Then	'Լ����ͻ
				strMsg = strMsg & "<br><li>������Ϣ�ѱ�ʹ�ã�</li>" & vbCrLf
			End If
			Err.Clear
			manySend = False
			Exit Function
		End If
		
		manySend = True
	End Function
	
	Function manyFinish()
		Dim dblId, bLock, arrInfoId, strUserSql
		Dim rsShop, strSql 

		dblId = trim(Request("shopId") & "")
		if dblId = "" then dblId = "0"
		
		Set rsShop = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select shopstate, infoId, FinishDate From shop_t Where shopstate = 2 and shopId in (" & dblId & ")"
		If rsShop.State = 1 Then rsShop.Close
		rsShop.Open strSql, conn, 1, 3
		
		If (rsShop.Bof Or rsShop.Eof) Then
			strMsg = strMsg & "<br><li>������Ϣ�����ڻ����Ѿ���ɾ����</li>" & vbCrLf
			If rsShop.State = 1 Then rsShop.Close
			Set rsShop = Nothing
			Exit Function
		else
			do while not(rsshop.bof or rsshop.eof)
				'���»�Ա�����ܴ���
				strUserSql = "Update info_t Set iCount = iCount+1 Where infoid = " & ConvertLong(rsshop("infoid") & "")
				conn.Execute strUserSql
				'���¶���״��
				rsshop("shopstate") = 3
				rsshop("FinishDate") = now()
				rsshop.update
				rsshop.movenext
			loop
		End IF
		
		If rsShop.State = 1 Then rsShop.Close
		Set rsShop = Nothing

		On Error Resume Next 

		If Err Then
			If Err.Number = -2147217900 Then	'Լ����ͻ
				strMsg = strMsg & "<br><li>������Ϣ�ѱ�ʹ�ã�</li>" & vbCrLf
			End If
			Err.Clear
			manyFinish = False
			Exit Function
		End If
		
		manyFinish = True
	End Function
	
	Function manyDelete()
		Dim dblId, arrInfoId, thisUserId
		Dim rsShop, strSql
		arrInfoId = "0"

		dblId = trim(Request("shopid") & "")
		thisUserId = ConvertLong(Request("userid") & "")
		if dblId = "" then dblId = "0"
		
		Set rsShop = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select shopstate, infoId From shop_t Where shopId in (" & dblId & ")"
		If rsShop.State = 1 Then rsShop.Close
		rsShop.Open strSql, conn, 1, 1
		
		If rsShop.Eof Then
			strMsg = strMsg & "<br><li>������Ϣ�����ڻ����Ѿ���ɾ����</li>" & vbCrLf
			If rsShop.State = 1 Then rsShop.Close
			Set rsShop = Nothing
			Exit Function
		else
			do while not(rsshop.bof or rsshop.eof)
				if ConvertLong(rsshop("shopstate") & "") < 3 then
					'�����鱾�����״̬
					strSql = "Update info_t Set iCount = iCount +1 Where infoId = " & ConvertLong(rsshop("infoId") & "")
					conn.Execute strSql
				end if
				rsshop.movenext
			loop
		End IF
		
		If rsShop.State = 1 Then rsShop.Close
		Set rsShop = Nothing

		On Error Resume Next
				
		'ɾ������
		strSql = "Delete From shop_t Where shopId in (" & dblId & ")"
		conn.Execute strSql

		If Err Then
			If Err.Number = -2147217900 Then	'Լ����ͻ
				strMsg = strMsg & "<br><li>��Ϣ�ѱ�ʹ�ã�</li>" & vbCrLf
			End If
			Err.Clear
			manyDelete = False
			Exit Function
		End If
		
		manyDelete = True
		
	End Function
%>
</body>
</html>
