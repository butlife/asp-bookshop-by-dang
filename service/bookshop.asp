<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<!--#include file="../common/safe.asp"-->
<%Response.ContentType = "application/json"%>
<%Response.Charset="UTF-8"%>
<%
	dim rsBook, strsql, iCount, strQuery
	dim lngSortId, strbookKeyWords, lngPageNum
	dim strTitle, lngInfoId, strPicUrl, lngisFav
	
	lngSortId = ConvertLong(request("SortId") & "")
	lngPageNum = ConvertLong(request("PageNum") & "")
	strbookKeyWords = trim(request("bookKeyWords") & "")
	
	strQuery = " where ispassed = 1 "
	if lngSortId <> 0 then
		strQuery = strQuery & " and sortid = " & lngSortId
	end if 
	
	if strbookKeyWords <> "" then
		strQuery = strQuery & " and (title like '%" & strbookKeyWords & "%' or content like '%" & strbookKeyWords & "%')"
	end if

	Set rsBook = Server.CreateObject("ADODB.RecordSet")
'	strsql = "select Title, InfoId, PicUrl from info_t " & strQuery & " order by istop desc, iorder desc"
	strsql = "select Title, InfoId, PicUrl, content from info_t " & strQuery & " order by infoid, title asc"
	if rsBook.state = 1 then rs.close
	rsBook.open strsql,conn,1,1
	rsBook.pagesize = glngPageSize_phone

	if lngPageNum > rsBook.PageCount then  
		rsBook.AbsolutePage = rsBook.PageCount  
	elseif lngPageNum <= 0 then  
		lngPageNum = 1
	else  
		rsBook.AbsolutePage = lngPageNum   
	end if  
	lngPageNum = rsBook.AbsolutePage	
%>
{
    "state": 0,
    "msg": "success",
	"data" :{
    	"sortid" : "<%=lngSortId%>",
        "bookKeyWords" : "<%=strbookKeyWords%>",
        "PageNum" : "<%=lngPageNum%>"
    },
    "body" : [
<%
	if not(rsBook.bof or rsBook.eof) then
		iCount=0
		do while not (rsBook.bof or rsBook.eof)
			lnginfoId = rsBook("infoId")
			strTitle = rsBook("Title")
			strpicurl = rsBook("picurl")
			strcontent = rsBook("content")
			lngisFav = getFav(lnginfoId)
%>
        {
            "infoId" : "<%=lnginfoId%>",
            "title" : "<%=strTitle%>",
            "picurl" : "<%=strpicurl%>",
            "content" : "<%=strcontent%>",
			"fav":"<%=lngisFav%>"
        }
<%
			iCount = iCount+1
			rsBook.movenext

            if not(rsBook.bof or rsBook.eof) and iCount < glngPageSize_phone then 
            'if not(rsBook.bof or rsBook.eof) then 
               response.write(",")
            end if

			if iCount >= glngPageSize_phone then exit do
		loop
	end if
%>
    ]
}
<%
	if rsBook.state = 1 then rsBook.close
	set rsBook = nothing
	
	function getFav(infoId)
		dim rsFav, strsql, strRe
		infoId = ConvertLong(infoId&"")
		Set rsFav = Server.CreateObject("ADODB.RecordSet")
		strsql = "select favId from fav_t where infoId = " & infoId & " and UserId = " & ConvertLong(Session("userid")&"")
		if rsFav.state = 1 then rsFav.close
		rsFav.open strsql,conn,1,1
		if not(rsFav.bof or rsFav.eof) then
			strRe = trim(rsFav("favId") & "")
		else
			strRe = "0"
		end if
		if rsFav.state = 1 then rsFav.close
		set rsFav = nothing
		getFav = strRe		
	end function

	
	Call CloseConn()
%>
