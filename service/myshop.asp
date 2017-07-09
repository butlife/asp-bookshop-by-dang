<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "application/json"%>
<%Response.Charset="UTF-8"%>
<%
	dim rsShop, strsql, iCount, strQuery
	dim lngUserId, strbookKeyWords, lngPageNum
	dim strTitle, lngShopId, strSendDate, lngshopState, strAddDate
	
	lngUserId = ConvertLong(session("UserId") & "")
	lngPageNum = ConvertLong(request("PageNum") & "")
	
	Set rsShop = Server.CreateObject("ADODB.RecordSet")
	strsql = "select * from shop_v where userid = " & lngUserId & " and ShopState < 3 order by shopId desc "
	if rsShop.state = 1 then rs.close
	rsShop.open strsql,conn,1,1
	rsShop.pagesize = glngPageSize

	pagecount = rsShop.PageCount

	if lngPageNum > rsShop.PageCount then  
		rsShop.AbsolutePage = rsShop.PageCount  
	elseif lngPageNum <= 0 then  
		lngPageNum = 1
	else  
		rsShop.AbsolutePage = lngPageNum   
	end if  
	lngPageNum = rsShop.AbsolutePage	
%>
{
    "state": 0,
    "msg": "success",
	"data" :{
    	"UserId" : "<%=lngUserId%>",
        "PageNum" : "<%=lngPageNum%>",
		"maxpagenum":<%=pagecount%>
    },
    "body" : [
<%
	if not(rsShop.bof or rsShop.eof) then
		iCount=0
		do while not (rsShop.bof or rsShop.eof)
			lngUserId = rsShop("UserId")
			lngShopId = rsShop("ShopId")
			strTitle = rsShop("Title")
			strSendDate = Format_Time(rsShop("sendDate"),2)
			strAddDate = Format_Time(rsShop("AddDate"),2)
			lngshopState = rsShop("shopState")
			strshopStateName = getState(lngshopState) 
%>
        {
            "UserId" : "<%=lngUserId%>",
            "ShopId" : "<%=lngShopId%>",
            "Title" : "<%=strTitle%>",
            "AddDate" : "<%=strAddDate%>",
            "SendDate" : "<%=strSendDate%>",
            "shopState" : "<%=lngshopState%>",
            "shopStateName" : "<%=strshopStateName%>"
        }
<%
			iCount = iCount+1
			rsShop.movenext

            if not(rsShop.bof or rsShop.eof) and iCount < glngPageSize then 
            'if not(rsShop.bof or rsShop.eof) then 
               response.write(",")
            end if

			if iCount >= glngPageSize then exit do
		loop
	end if
%>
    ]
}
<%
	if rsShop.state = 1 then rsShop.close
	set rsShop = nothing

	function getState(i)
		dim strRe
		i = trim(i & "")
		select case i
			case "0"
			strRe = "己下单"
			
			case "1"
			strRe = "己送货"
			
			case "2"
			strRe = "己归还"
			
			case "3"
			strRe = "己完结"
			
			case else
			strRe = ""
		end select
		getState = strRe
	end function
	
	Call CloseConn()
%>
