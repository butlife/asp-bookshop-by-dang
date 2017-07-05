<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "text/json"%>
<%
'插入会员收藏书本，传入书本ID，读取session 中的userID
	dim lngUserId, strArrShopId, i
	dim rsShop, strsql, lngstate, strMsg
	
	lngUserId = ConvertLong(Session("UserId") & "")
	strArrShopId = trim(request("ShopId") & "")
	if strArrShopId = "" then strArrShopId = 0
	
	Set rsShop = Server.CreateObject("ADODB.RecordSet")
	strsql = "select * from shop_v where shopstate = 1 and UserId = " & lngUserId & " and ShopId in (" & strArrShopId & ")"
	'response.write strsql
	if rsShop.state = 1 then rs.close
	rsShop.open strsql,conn,1,3
	if not(rsShop.bof or rsShop.eof) then
		i=0
		do while not(rsShop.bof or rsShop.eof)
			rsShop("shopstate") = 2
			'rsShop("icount") = rsShop("icount")+1
			rsShop("returnedDate") = now()
			i=i+1
			rsShop.update
			rsshop.movenext
		loop
		lngstate = 0
		strMsg = i & "本书，申请归还成功"
	else
		lngstate = 1
		strMsg = "请选中要归还的书本！"
	end if
	if rsShop.state = 1 then rsShop.close
	set rsShop = nothing
	
	Call CloseConn()
%>
{
    "state": "<%=lngstate%>",
    "msg": "<%=strMsg%>"
}