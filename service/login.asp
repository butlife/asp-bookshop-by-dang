<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<%Response.ContentType = "text/json"%>
<%
	dim strUserAcc, strUserPwd
	dim rsUser, strsql, lngstate, strMsg, strIP
	dim bispassed, strexpdate_s, strexpdate_e, lngUserId
	
	strUserAcc = trim(request("UserAcc") & "")
	strUserPwd = trim(request("UserPwd") & "")
	
	strIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If strIP = "" Then
		strIP = Request.ServerVariables("REMOTE_ADDR") 
	end if
	if strIP = "::1" or strIP = "" then strIP = "未知IP"

		
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	strsql = "select userId, ispassed, expdate_s, expdate_e, loginCounts, loginDate, loginIP from user_t where useracc = '" & strUserAcc & "' and userpwd = '" & strUserPwd & "'"
	if rsUser.state = 1 then rs.close
	rsuser.open strsql,conn,1,3
	if not(rsUser.bof or rsUser.eof) then
		lngUserId = ConvertLong(rsuser("UserId") & "")
		bispassed = ConvertLong(rsuser("ispassed") & "")
		strexpdate_s = formatdatetime(rsuser("expdate_s"),1)
		strexpdate_e = formatdatetime(rsuser("expdate_e"),1)
		if (bispassed = 1) then
			if (DateDiff ("d", strexpdate_s, date()) < 0 or DateDiff("d", strexpdate_e, date()) >0) then
				lngstate = 1
				strMsg = "登录失败，帐号己过期！"
			else
				lngstate = 0
				strMsg = "登录成功"
				rsuser("loginCounts") = rsuser("loginCounts") + 1
				rsUser("loginDate") = now()
				rsUser("loginIP") = strIP
				rsUser.update
				Session("userid") = lngUserId
				Session("useracc") = strUserAcc
			end if
		else
			lngstate = 1
			strMsg = "登录失败，帐号未审核！"
		end if
	else
		lngstate = 1
		strMsg = "登录失败，帐号或密码错误！"
	end if
	if rsUser.state = 1 then rsUser.close
	set rsUser = nothing
	
	Call CloseConn()
%>
{
    "state": "<%=lngstate%>",
    "msg": "<%=strMsg%>"
}