<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "application/json"%>
<%Response.Charset="UTF-8"%>
<%
'插入会员收藏书本，传入书本ID，读取session 中的userID
	dim lngUserId, lngInfoId, lngFavState
	dim rsFav, strsql, lngstate, strMsg
	
	lngUserId = ConvertLong(Session("UserId") & "")
	lngInfoId = ConvertLong(request("InfoId") & "")
	lngFavState = ConvertLong(request("fav") & "")
		
	Set rsFav = Server.CreateObject("ADODB.RecordSet")
	
	if lngFavState = 0 then
		strsql = "select * from Fav_t where UserId = " & lngUserId & " and InfoId = " & lngInfoId
		if rsFav.state = 1 then rs.close
		rsFav.open strsql,conn,1,3
		if (rsFav.bof or rsFav.eof) then
			rsFav.addnew
			rsFav("UserId") = lngUserId
			rsFav("InfoId") = lngInfoId
			rsFav("Favdate") = now()
			rsFav.update
			lngstate = 0
			strMsg = "收藏成功"
		else
			lngstate = 1
			strMsg = "己收藏"
		end if
	else
		strsql = "Delete From Fav_t where UserId = " & lngUserId & " and InfoId = " & lngInfoId
		conn.Execute strsql
		
		lngstate = 2
		strMsg = "己取消收藏"
	end if
	if rsFav.state = 1 then rsFav.close
	set rsFav = nothing
	
	Call CloseConn()
%>
{
    "state": "<%=lngstate%>",
    "msg": "<%=strMsg%>"
}