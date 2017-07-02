<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<%'Response.ContentType = "text/json"%>
<%
	dim rsFav, strsql, iCount, strQuery
	dim lngUserId, strbookKeyWords, lngPageNum
	dim strTitle, lngInfoId, strPicUrl
	
	lngUserId = ConvertLong(session("UserId") & "")
	lngPageNum = ConvertLong(request("PageNum") & "")
	
'	strQuery = " where ispassed = 1 "
'	if lngSortId <> 0 then
'		strQuery = strQuery & " and sortid = " & lngSortId
'	end if 
'	
'	if strbookKeyWords <> "" then
'		strQuery = strQuery & " and (title like '%" & strbookKeyWords & "%' or content like '%" & strbookKeyWords & "%')"
'	end if

	Set rsFav = Server.CreateObject("ADODB.RecordSet")
'	strsql = "select Title, InfoId, PicUrl from info_t " & strQuery & " order by istop desc, iorder desc"
	strsql = "select * from fav_t order by favid desc "
	if rsFav.state = 1 then rs.close
	rsFav.open strsql,conn,1,1
	rsFav.pagesize = glngPageSize_phone

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
        "PageNum" : "<%=lngPageNum%>"
    },
    "body" : [
<%
	if not(rsFav.bof or rsFav.eof) then
		iCount=0
		do while not (rsFav.bof or rsFav.eof)
			lngUserId = rsFav("UserId")
			lngInfoId = rsFav("infoId")
			strTitle = getBookTitle(lngInfoId)
			strFavDate = rsFav("FavDate")
%>
        {
            "UserId" : "<%=lngUserId%>",
            "infoId" : "<%=lngInfoId%>",
            "Title" : "<%=strTitle%>",
            "FavDate" : "<%=strFavDate%>"
        }
<%
			iCount = iCount+1
			rsFav.movenext

            if not(rsFav.bof or rsFav.eof) and iCount < glngPageSize_phone then 
            'if not(rsFav.bof or rsFav.eof) then 
               response.write(",")
            end if

			if iCount >= glngPageSize_phone then exit do
		loop
	end if
%>
    ]
}
<%
	if rsFav.state = 1 then rsFav.close
	set rsFav = nothing
	
	Call CloseConn()
	
	function getBookTitle(Id)
		dim rsBook, strsortSql, strRe
		strRe = "书名"
		Set rsBook = Server.CreateObject("ADODB.RecordSet")
		Id = ConvertLong(Id & "")
		strBookSql = "Select Title From info_t where infoId =  " & Id
		rsBook.Open strBookSql, conn, 1, 1
		if not(rsBook.bof or rsBook.eof) then
			strRe = trim(rsBook("Title")&"")
		end if
		If rsBook.State = 1 Then rsBook.Close
		Set rsBook = Nothing
		getBookTitle = strRe	
	end function
%>
