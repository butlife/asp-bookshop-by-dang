﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "application/json"%>
<%Response.Charset="UTF-8"%>
<%
	dim rsFav, strsql, iCount, strQuery
	dim lngUserId, strbookKeyWords, lngPageNum
	dim strTitle, lngInfoId, strPicUrl, lngFavId
	
	lngUserId = ConvertLong(session("UserId") & "")
	lngPageNum = ConvertLong(request("PageNum") & "")

	Set rsFav = Server.CreateObject("ADODB.RecordSet")
	strsql = "select * from fav_v where userId = " & lngUserId & " and ispassed = 1 order by favid desc "
	if rsFav.state = 1 then rs.close
	rsFav.open strsql,conn,1,1
	rsFav.pagesize = glngPageSize

	pagecount = rsFav.PageCount

	if lngPageNum > rsFav.PageCount then  
		rsFav.AbsolutePage = rsFav.PageCount  
	elseif lngPageNum <= 0 then  
		lngPageNum = 1
	else  
		rsFav.AbsolutePage = lngPageNum   
	end if  
	lngPageNum = rsFav.AbsolutePage	
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
	if not(rsFav.bof or rsFav.eof) then
		iCount=0
		do while not (rsFav.bof or rsFav.eof)
			lngFavId = rsFav("FavId")
			lngInfoId = rsFav("infoId")
			strTitle = rsFav("Title")
			strFavDate = Format_Time(rsFav("FavDate"),2)
			lngiCount = rsFav("iCount")
%>
        {
            "FavId" : "<%=lngFavId%>",
            "infoId" : "<%=lngInfoId%>",
            "Title" : "<%=strTitle%>",
            "iCount" : "<%=lngiCount%>",
            "FavDate" : "<%=strFavDate%>"
        }
<%
			iCount = iCount+1
			rsFav.movenext

            if not(rsFav.bof or rsFav.eof) and iCount < glngPageSize then 
            'if not(rsFav.bof or rsFav.eof) then 
               response.write(",")
            end if

			if iCount >= glngPageSize then exit do
		loop
	end if
%>
    ]
}
<%
	if rsFav.state = 1 then rsFav.close
	set rsFav = nothing
	
	Call CloseConn()
%>
