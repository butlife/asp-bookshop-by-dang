<%
	'============================================
	'��������WriteErrMsg
	'��  �ã���ʾ������ʾ��Ϣ
	'��  ����strErrMsg   ������ʾ��Ϣ
	'����ֵ����
	'============================================
	Sub WriteErrorMsg(strErrMsg)
		Dim strErr, ComeUrl
		ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
		
		strErr = strErr & "<html><head><title>������Ϣ</title><meta http-equiv = 'Content-Type' content = 'text/html; charset = gb2312'>" & vbcrlf
		strErr = strErr & "<link href = '" & gstrInstallDir & "CSS/style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
		strErr = strErr & "<table cellpadding = 0 cellspacing = 0 cellpadding='0' border='1'cellspacing='0' bordercolor='#CCCCCC' bordercolorlight='#CCCCCC' width = 400 style='margin-top:40px;' align = center>" & vbcrlf
		strErr = strErr & "  <tr align = 'center' style = 'background-color:#CCCCCC; color:red;'><td height = '25' align=left><span><img src='" & gstrInstallDir & "Images/error.gif' heigh='16' width='16'>&nbsp;������Ϣ</span></td></tr>" & vbcrlf
		strErr = strErr & "  <tr style = 'background-color:#EEEEEE;'><td height = '100' valign = 'top' style='padding:10px;'><b>��������Ŀ���ԭ��</b>" & strErrMsg &"</td></tr>" & vbcrlf
		strErr = strErr & "  <tr align = 'center'style='height:25px; background-color:#cccccc;'><td>"
		If ComeUrl <> "" then
			strErr = strErr & "<script language='javascript'>window.setTimeout(""window.history.back();"", 2000);</script>"
'			strErr = strErr & "<a href = '" & ComeUrl & "'>&lt;&lt; ������һҳ</a>"
			strErr = strErr & "<a href = 'javascript:history.back();'>&lt;&lt; ������һҳ</a>"
		Else
			strErr = strErr & "<a href = 'javascript:window.close();'>���رա�</a>"
		End If
		strErr = strErr & "</td></tr>" & vbcrlf
		strErr = strErr & "</table>" & vbcrlf
		strErr = strErr & "</body></html>" & vbcrlf
		Response.Write strErr
	End Sub
	
	'============================================
	'��������WriteSuccessMsg
	'��  �ã���ʾ�ɹ���ʾ��Ϣ
	'��  ����strSuccessMsg   �ɹ���ʾ��Ϣ
	'		 ComeUrl         ����ҳ��   
	'����ֵ����
	'============================================
	Sub WriteSuccessMsg(strSuccessMsg, GotoUrl)
		Dim strSuccess, ComeUrl, bGoBack
		ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
		If Trim(GotoUrl) <> "" Then
			If GotoUrl = "False" Or GotoUrl = "True" Then
				bGoBack = False
				GotoUrl = ""
			End If
			ComeUrl = GotoUrl
		End If
		
		strSuccess = strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv = 'Content-Type' content = 'text/html; charset = gb2312'>" & vbcrlf
		strSuccess = strSuccess & "<link href = '" & gstrInstallDir & "Css/Style.css' rel = 'stylesheet' type = 'text/css'></head><body><br><br>" & vbcrlf
		strSuccess = strSuccess & "<table width = 400 border='1' cellpadding='0' cellspacing='0' bordercolor='#cccccc' bordercolorlight='#cccccc' align = center style='margin-top:40px;'>" & vbcrlf
		strSuccess = strSuccess & "  <tr align = 'Left' style = 'background-color:#CCCCCC; color:red;'><td height = '25'>&nbsp;&nbsp;��ϲ�㣡</td></tr>" & vbcrlf
		strSuccess = strSuccess & "  <tr style = 'background-color:#EEEEEE;'><td height = '100' valign = 'top' style='padding:10px;'>" & strSuccessMsg &"</td></tr>" & vbcrlf
		strSuccess = strSuccess & "  <tr align = 'center' style = 'background-color:#CCCCCC; height:25px;'><td>"
		If ComeUrl <> "" then
			strSuccess = strSuccess & "<script language='javascript'>window.setTimeout(""window.location.href='" & ComeUrl & "';"", 1000);</script>"
			strSuccess = strSuccess & "<a href = '" & ComeUrl & "'>&lt;&lt; ������һҳ</a>"
		else
			strSuccess = strSuccess & "<a href = 'javascript:window.close();'>���رա�</a>"
		End If
		strSuccess = strSuccess & "</td></tr>" & vbcrlf
		strSuccess = strSuccess & "</table>" & vbcrlf
		strSuccess = strSuccess & "</body></html>" & vbcrlf
		Response.Write strSuccess
	End Sub
	
	'============================================
	'��������ConfirmBox
	'��  �ã���ʾȷ����ʾ��Ϣ
	'��  ����strConfirmMsg   ȷ����ʾ��Ϣ
	'����ֵ����
	'============================================
	Sub ConfirmBox(strConfirmMsg)
		Dim strHtml

		strHtml = ""
		strHtml = strHtml & "<script language=""javascript"">" & vbCrLf
		strHtml = strHtml & "<!--" & vbCrLf
		strHtml = strHtml & "temp=window.confirm(""" & Trim(strConfirmMsg) & """);" & vbCrLf
		strHtml = strHtml & "if (temp){" & vbCrLf
		strHtml = strHtml & "}" & vbCrLf
		strHtml = strHtml & "else{" & vbCrLf
		strHtml = strHtml & "history.back();" & vbCrLf
		strHtml = strHtml & "}" & vbCrLf
		strHtml = strHtml & "//-->" & vbCrLf
		strHtml = strHtml & "</SCRIPT>" & vbCrLf
		Response.Write strHtml
		
	End Sub
%>