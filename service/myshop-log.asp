<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "application/json"%>
<%Response.Charset="UTF-8"%>
<%
	dim rsShop, strsql, iCount, strQuery
	dim lngUserId, strbookKeyWords, lngPageNum
	dim strTitle, lngShopId, strreturnedDate
	
	lngUserId = ConvertLong(session("UserId") & "")
	lngPageNum = ConvertLong(request("PageNum") & "")
	
	Set rsShop = Server.CreateObject("ADODB.RecordSet")
	strsql = "select * from shop_v where userid = " & lngUserId & " and ShopState > 1 order by shopId desc "
	if rsShop.state = 1 then rs.close
	rsShop.open strsql,conn,1,1
	rsShop.pagesize = glngPageSize

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
    	"ShopState" : "1",
        "PageNum" : "<%=lngPageNum%>"
    },
    "body" : [
<%
	if not(rsShop.bof or rsShop.eof) then
		iCount=0
		do while not (rsShop.bof or rsShop.eof)
			lngUserId = rsShop("UserId")
			lngShopId = rsShop("ShopId")
			strTitle = rsShop("Title")
			strreturnedDate = Format_Time(rsShop("returnedDate"),2)
%>
        {
            "iCount" : "<%=iCount%>",
            "UserId" : "<%=lngUserId%>",
            "ShopId" : "<%=lngShopId%>",
            "Title" : "<%=strTitle%>",
            "returnedDate" : "<%=strreturnedDate%>"
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
	
	Call CloseConn()
%>
