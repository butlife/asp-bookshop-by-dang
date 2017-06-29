<%
	'============================================
	'过程名：WriteErrMsg
	'作  用：显示错误提示信息
	'参  数：strErrMsg   错误提示信息
	'返回值：无
	'============================================
	Sub WriteErrorMsg(strErrMsg)
		Dim strErr, ComeUrl
		ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
		
		strErr = strErr & "<html><head><title>错误信息</title><meta http-equiv = 'Content-Type' content = 'text/html; charset = gb2312'>" & vbcrlf
		strErr = strErr & "<link href = '" & gstrInstallDir & "CSS/style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
		strErr = strErr & "<table cellpadding = 0 cellspacing = 0 cellpadding='0' border='1'cellspacing='0' bordercolor='#CCCCCC' bordercolorlight='#CCCCCC' width = 400 style='margin-top:40px;' align = center>" & vbcrlf
		strErr = strErr & "  <tr align = 'center' style = 'background-color:#CCCCCC; color:red;'><td height = '25' align=left><span><img src='" & gstrInstallDir & "Images/error.gif' heigh='16' width='16'>&nbsp;错误信息</span></td></tr>" & vbcrlf
		strErr = strErr & "  <tr style = 'background-color:#EEEEEE;'><td height = '100' valign = 'top' style='padding:10px;'><b>产生错误的可能原因：</b>" & strErrMsg &"</td></tr>" & vbcrlf
		strErr = strErr & "  <tr align = 'center'style='height:25px; background-color:#cccccc;'><td>"
		If ComeUrl <> "" then
			strErr = strErr & "<script language='javascript'>window.setTimeout(""window.history.back();"", 2000);</script>"
'			strErr = strErr & "<a href = '" & ComeUrl & "'>&lt;&lt; 返回上一页</a>"
			strErr = strErr & "<a href = 'javascript:history.back();'>&lt;&lt; 返回上一页</a>"
		Else
			strErr = strErr & "<a href = 'javascript:window.close();'>【关闭】</a>"
		End If
		strErr = strErr & "</td></tr>" & vbcrlf
		strErr = strErr & "</table>" & vbcrlf
		strErr = strErr & "</body></html>" & vbcrlf
		Response.Write strErr
	End Sub
	
	'============================================
	'过程名：WriteSuccessMsg
	'作  用：显示成功提示信息
	'参  数：strSuccessMsg   成功提示信息
	'		 ComeUrl         调用页面   
	'返回值：无
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
		
		strSuccess = strSuccess & "<html><head><title>成功信息</title><meta http-equiv = 'Content-Type' content = 'text/html; charset = gb2312'>" & vbcrlf
		strSuccess = strSuccess & "<link href = '" & gstrInstallDir & "Css/Style.css' rel = 'stylesheet' type = 'text/css'></head><body><br><br>" & vbcrlf
		strSuccess = strSuccess & "<table width = 400 border='1' cellpadding='0' cellspacing='0' bordercolor='#cccccc' bordercolorlight='#cccccc' align = center style='margin-top:40px;'>" & vbcrlf
		strSuccess = strSuccess & "  <tr align = 'Left' style = 'background-color:#CCCCCC; color:red;'><td height = '25'>&nbsp;&nbsp;恭喜你！</td></tr>" & vbcrlf
		strSuccess = strSuccess & "  <tr style = 'background-color:#EEEEEE;'><td height = '100' valign = 'top' style='padding:10px;'>" & strSuccessMsg &"</td></tr>" & vbcrlf
		strSuccess = strSuccess & "  <tr align = 'center' style = 'background-color:#CCCCCC; height:25px;'><td>"
		If ComeUrl <> "" then
			strSuccess = strSuccess & "<script language='javascript'>window.setTimeout(""window.location.href='" & ComeUrl & "';"", 1000);</script>"
			strSuccess = strSuccess & "<a href = '" & ComeUrl & "'>&lt;&lt; 返回上一页</a>"
		else
			strSuccess = strSuccess & "<a href = 'javascript:window.close();'>【关闭】</a>"
		End If
		strSuccess = strSuccess & "</td></tr>" & vbcrlf
		strSuccess = strSuccess & "</table>" & vbcrlf
		strSuccess = strSuccess & "</body></html>" & vbcrlf
		Response.Write strSuccess
	End Sub
	
	'============================================
	'过程名：ConfirmBox
	'作  用：显示确认提示信息
	'参  数：strConfirmMsg   确认提示信息
	'返回值：无
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