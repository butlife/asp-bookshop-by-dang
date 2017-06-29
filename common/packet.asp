<%
	function newslist(iSortId,iCount, iLen)
		iSortId = ConvertLong(iSortId & "")
		iCount = ConvertLong(iCount & "")
		iLen = ConvertLong(iLen & "")
		Dim strsql, strtitle, dblId, dtmakedate, rs, i
		set rs = server.CreateObject("adodb.recordset")
		'strsql = " select infoId, title, makedate from info_t where ispassed = 1 and istop = 1 order by istop desc, iorder desc, makedate desc, infoid desc"
		strsql = " select infoId, title, makedate from info_t where ispassed = 1 and sortid = " & iSortId & " order by istop desc, iorder desc, makedate desc, infoid desc"
		if rs.state = 1 then rs.close
		rs.open strsql, conn,1,1
		response.write "<table border=""0"" align=""center"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
		for i = 1 to iCount
			if not(rs.bof or rs.eof) then
				dblId = ConvertLong(rs("infoId") &"")
				strtitle = mid(trim(rs("title") &""),1,iLen)
				dtmakedate = Format_Time(rs("makedate"),9)
				response.write "<tr><td class=""indexnews""><a href=""news.asp?id=" & dblId & """>" & strtitle & "</a></td><td style=""width:40px;"">" & dtmakedate & "</td></tr>"
				rs.movenext
			else
				response.write "<tr><td class=""indexnews"">&nbsp;</td><td>&nbsp;</td></tr>"
			end if
		next
		response.write "</table>"
		if rs.state = 1 then rs.close
		set rs = nothing
	end function
	
	function job(iCount, iLen)
		iSortId = ConvertLong(iSortId & "")
		iCount = ConvertLong(iCount & "")
		iLen = ConvertLong(iLen & "")
		Dim strsql, strtitle, dblId, dtmakedate, rs, i
		set rs = server.CreateObject("adodb.recordset")
		'strsql = " select infoId, title, makedate from info_t where ispassed = 1 and istop = 1 order by istop desc, iorder desc, makedate desc, infoid desc"
		strsql = " select infoId, title, makedate from info_t where ispassed = 1 and sortid in(26,28) order by istop desc, iorder desc, makedate desc, infoid desc"
		if rs.state = 1 then rs.close
		rs.open strsql, conn,1,1
		response.write "<table border=""0"" align=""center"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
		for i = 1 to iCount
			if not(rs.bof or rs.eof) then
				dblId = ConvertLong(rs("infoId") &"")
				strtitle = mid(trim(rs("title") &""),1,iLen)
				dtmakedate = Format_Time(rs("makedate"),9)
				response.write "<tr><td class=""indexnews""><a href=""news.asp?id=" & dblId & """>" & strtitle & "</a></td><td style=""width:40px;"">" & dtmakedate & "</td></tr>"
				rs.movenext
			else
				response.write "<tr><td class=""indexnews"">&nbsp;</td><td>&nbsp;</td></tr>"
			end if
		next
		response.write "</table>"
		if rs.state = 1 then rs.close
		set rs = nothing
	end function
	
	function gg(iLen)
		iSortId = ConvertLong(iSortId & "")
		iCount = ConvertLong(iCount & "")
		iLen = ConvertLong(iLen & "")
		Dim strsql, strtitle, dblId, dtmakedate, rs, i
		set rs = server.CreateObject("adodb.recordset")
		strsql = " select top 1 infoId, title, makedate,content from info_t where ispassed = 1 and istop = 1 and sortid = 25 order by istop desc, iorder desc, makedate desc, infoid desc"
		if rs.state = 1 then rs.close
		rs.open strsql, conn,1,1
			if not(rs.bof or rs.eof) then
				dblId = ConvertLong(rs("infoId") &"")
				'strtitle = mid(trim(rs("title") &""),1,iLen)
				strGGBody = mid(CleanHTML(trim(rs("content") &"")),1,iLen)
				'dtmakedate = Format_Time(rs("makedate"),9)
				response.write "<a href=""news.asp?id=" & dblId & """>" & strGGBody & "</a>"
				rs.movenext
			else
				response.write "&nbsp;"
			end if
		if rs.state = 1 then rs.close
		set rs = nothing
	end function

	function about(Id, iCount)
		dim strAbout
		Id = ConvertLong(Id & "")
		set rs = server.CreateObject("adodb.recordset")
		sql = " select * from about_t where id = " & Id
		if rs.state = 1 then rs.close
		rs.open sql, conn, 1, 1
		If (rs.bof or rs.eof) then
			Response.write "&nbsp;"
		else
			strAbout = RemoveHTML(rs("content")&"")
			Response.write mid(strAbout,1,iCount) & "..."
		end if
		if rs.state = 1 then rs.close
		set rs = nothing
	end function
	
	function pro_nav(iCount)
		iCount = ConvertLong(iCount & "")
		dim strsql, rs
		set rs = server.CreateObject("adodb.recordset")
		sql = " select * from product_t where ispassed = 1 order by istop desc, iorder desc, productId "
		if rs.state = 1 then rs.close
		rs.open sql, conn, 1, 1
		if not(rs.bof or rs.eof) then
			i = 0
			do while not(rs.bof or rs.eof) and (i<iCount)
				response.write "<tr><td><a href=""product.asp?id=" & rs("productId") & """> " & rs("title") & "</a></td></tr>"
				i = i + 1
				rs.movenext
			loop
		end if
		if rs.state = 1 then rs.close
		set rs = nothing
	end function
%>