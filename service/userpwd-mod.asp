﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "application/json"%>
<%Response.Charset="UTF-8"%>
<%
	dim strUserAcc, strchkuserpwd, lngUserId, struserpwd
	dim rsUser, strsql, lngstate, strMsg
	dim strnewuserpwd, strchknewuserpwd
	
	lngUserId = ConvertLong(Session("UserId") & "")
	strUserAcc = trim(Session("UserAcc") & "")
	strchkuserpwd = trim(request("chkuserpwd") & "")
	strnewuserpwd = trim(request("newuserpwd") & "")
	strchknewuserpwd = trim(request("chknewuserpwd") & "")
		
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	strsql = "select userpwd from user_t where useracc = '" & strUserAcc & "' and userid = " & lngUserId
	if rsUser.state = 1 then rs.close
	rsuser.open strsql,conn,1,3
	if not(rsUser.bof or rsUser.eof) then
		struserpwd = trim(rsuser("userpwd") & "")
		if (strchkuserpwd = struserpwd) then
				rsuser("userpwd") = strnewuserpwd
				rsUser.update
				lngstate = 0
				strMsg = "修改资料成功"
		else
			lngstate = 1
			strMsg = "更新失败，密码错误！"
		end if
	else
		lngstate = 1
		strMsg = "更新失败，请重新登录后再试！"
	end if
	if rsUser.state = 1 then rsUser.close
	set rsUser = nothing
	
	Call CloseConn()
%>
{
    "state": "<%=lngstate%>",
    "msg": "<%=strMsg%>"
}