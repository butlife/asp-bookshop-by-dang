<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "text/json"%>
<%
'�����Ա�ղ��鱾�������鱾ID����ȡsession �е�userID
	dim lngUserId, strArrFavId, i
	dim rsFav, strsql, lngstate, strMsg, strDelSql
	
	lngUserId = ConvertLong(Session("UserId") & "")
	strArrFavId = trim(request("FavId") & "")
	if strArrFavId = "" then strArrFavId = 0
	
	Set rsFav = Server.CreateObject("ADODB.RecordSet")
	strsql = "select FavId from Fav_t where UserId = " & lngUserId & " and FavId in (" & strArrFavId & ")"
	if rsFav.state = 1 then rs.close
	rsFav.open strsql,conn,1,1
	if not(rsFav.bof or rsFav.eof) then
		i=0
		do while not(rsFav.bof or rsFav.eof)
			strDelSql = "Delete From Fav_t Where FavId = " & ConvertLong(rsFav("FavId") & "")
			conn.Execute strDelSql
			i=i+1
			rsFav.movenext
		loop
		lngstate = 0
		strMsg = i & "����¼ȡ���ղسɹ�"
	else
		lngstate = 1
		strMsg = "��ѡ��Ҫȡ���ղص��鱾��"
	end if
	if rsFav.state = 1 then rsFav.close
	set rsFav = nothing
	
	Call CloseConn()
%>
{
    "state": "<%=lngstate%>",
    "msg": "<%=strMsg%>"
}