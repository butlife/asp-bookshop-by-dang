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
	if lngPageNum = 0 then lngPageNum = 1
	
	strQuery = " where ispassed = 1 "
	if lngSortId <> 0 then
		strQuery = strQuery & " and sortid = " & lngSortId
	end if 
	
	if strbookKeyWords <> "" then
		strQuery = strQuery & " and (title like '%" & strbookKeyWords & "%' or content like '%" & strbookKeyWords & "%')"
	end if
	%>
{
    "state": 0,
    "msg": "ok",
	"data" :{
    	"sortid" : "<%=lngSortId%>",
        "bookKeyWords" : "<%=strbookKeyWords%>",
        "PageNum" : "<%=lngPageNum%>"
    },
    "body" : [
<%
	Set rsBook = Server.CreateObject("ADODB.RecordSet")
	strsql = "select Title, InfoId, PicUrl from info_t " & strQuery & " order by istop desc, iorder desc"
	if rsBook.state = 1 then rs.close
	rsBook.open strsql,conn,1,1
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
            "picurl" : "<%=strpicurl%>"
        }
<%
			iCount = iCount+1
			rsBook.movenext

            if not(rsBook.bof or rsBook.eof) and iCount < 10 then 
            'if not(rsBook.bof or rsBook.eof) then 
               response.write(",")
            end if

			if iCount >= 10 then exit do
		loop
	end if
%>
    ]
}
<%
	if rsBook.state = 1 then rsBook.close
	set rsBook = nothing
%>
