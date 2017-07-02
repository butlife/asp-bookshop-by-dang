<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../common/conn-utf.asp"-->
<!--#include file="../common/Function-utf.asp"-->
<%Response.ContentType = "text/json"%>
<%
	dim rsBook, strsql, iCount, strQuery
	dim lngSortId, strbookKeyWords, lngPageNum
	dim strTitle, lngInfoId, strPicUrl
	
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
	strsql = "select Title, InfoId, PicUrl from info_t " & strQuery & " order by infoid, title asc"
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
%>
        {
            "infoId" : "<%=lnginfoId%>",
            "title" : "<%=strTitle%>",
            "picurl" : "<%=strpicurl%>",
			"fav":"0"
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
	
	Call CloseConn()
%>
