<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/Pagination.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>js/common.js" type="text/javascript"></script>
<script language="javascript" src="<%= gstrInstallDir%>js/checkbox.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
function Body_Load(){
}
</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">��Ϣ�б�</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">��Ϣ���</a>
    <a class="aButton" style="cursor:pointer;" onClick="AllDelete('CHK');">����ɾ��</a>
<%
	Dim lngInfoID, strtitle, bispassed, dtmakedate, strsortname, stradminname, ihit, strauthor, bistop, iorder, strKeyWords, strpicurl
	Dim rsInfo, i, strSql, strQuery, lngSortId
	Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
	
	lngSortId = ConvertDouble(Request("SortId") & "")
	strQuery = Trim(Request("Query") & "")
	
	sortlist(lngSortId)
	
	Function sortlist(sortid)
		sortid = ConvertDouble(sortid & "")
		Dim rsSort, strSql, strsortname, lngsortid
		set rssort = server.CreateObject("adodb.recordset")
		strsql = " select * from infosort_t "
		if rssort.state = 1 then rssort.close
		rssort.open strsql, conn, 1,1
		if not(rssort.bof or rssort.eof) then
			response.write "<a href='main.asp'"
			if (sortid = 0) Then
				response.write " class=""main_sort_red_cls"""
			end if
				response.write ">ȫ��"
			response.write "</a> &nbsp;&nbsp;"
			do while not(rssort.bof or rssort.eof)
				response.write "<a href='main.asp?sortid=" & ConvertDouble(rssort("sortid")& "") & "' "
				if (sortid = ConvertDouble(rssort("sortid")& "")) Then
					response.write " class=""main_sort_red_cls"""
				end if
				response.write ">"
				response.write trim(rssort("sortname")& "")
				response.write "</a> &nbsp;&nbsp;"
				rssort.movenext
			loop
		Else
			response.write "�������"
		end if
	end function
%>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" width="98%" cellspacing="0">
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap" width="35">&nbsp;</th>
            <th nowrap="nowrap"><input name="MAINCHK" id="MAINCHK" type="checkbox" value="" onClick="cheageBox('MAINCHK','CHK');"></th>
            <th nowrap="nowrap">����</th>
            <th nowrap="nowrap">���</th>
            <th nowrap="nowrap">���|�ö�</th>
            <th nowrap="nowrap">ͼƬ</th>
            <th nowrap="nowrap">����ʱ��</th>
            <th nowrap="nowrap">����</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		
		If Trim(strQuery) = "" Then
			lngSortId = ConvertDouble(Request("SortId") & "")		
			
			strQuery = " Where 1 = 1 "
				If lngSortId <> 0 Then
					strQuery = strQuery & " And SortId = " & lngSortId
				End If
		Else
			strQuery = outHTML(strQuery)
		End If
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From info_v " & strQuery & " Order By istop desc, ispassed desc, iorder desc, infoid desc "
		rsInfo.Open strSql, conn, 1, 1
		
		If Not (rsInfo.Bof Or rsInfo.Eof) Then
			bPagination = True
			'��ҳ
			lngPageSize = glngPageSize
			lngRecordCount = rsInfo.RecordCount
			rsInfo.PageSize = lngPageSize
			lngPageCount = rsInfo.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsInfo.AbsolutePage = lngCurrPage
			
			i = 0
			
			Do While Not (rsInfo.Bof Or rsInfo.Eof) 

				lngInfoID = ConvertDouble(rsInfo("InfoID") & "")
				strtitle = trim(rsInfo("title") & "")
				bispassed = ConvertDouble(rsInfo("ispassed") & "")
				bistop = ConvertDouble(rsInfo("istop") & "")
				dtmakedate = Format_Time(rsInfo("makedate"),7)
				strsortname = Trim(rsInfo("sortname") & "")
				stradminname = Trim(rsInfo("adminname") & "")
				ihit = ConvertDouble(rsInfo("hit") & "")
				'strauthor = Trim(rsInfo("author") & "")
				iorder = ConvertDouble(rsInfo("iorder") & "")
				'strKeyWords = Trim(rsInfo("KeyWords") & "")
				strpicurl = Trim(rsInfo("picurl") & "")
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center""><input name=""CHK"" id=""CHK"" type=""checkbox"" value=""" & lngInfoId & """></td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""left"">" & TdString(strtitle) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strsortname) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & getState(bispassed) & "|" & getState(bistop) & "</td>" & vbCrLf
				if strpicurl = "" then
					Response.Write "<td nowrap=""nowrap"" align=""center"">&nbsp;</td>" & vbCrLf
				else
					Response.Write "<td nowrap=""nowrap"" align=""center""><a href=""../../uppic/big/" & strpicurl & """ target=""_blank"">ͼƬ</a></td>" & vbCrLf
				end if
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(dtmakedate) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngInfoID & "'>�޸�</a>" & vbCrLf
				If (bistop = 1) Then
					Response.Write "<a class='admin_btn_b' href='action.asp?type=top&tbId=" & lngInfoID & "'>���</a>" & vbCrLf
				Else
					Response.Write "<a class='admin_btn' href='action.asp?type=top&tbId=" & lngInfoID & "'>�ö�</a>" & vbCrLf
				End If
				Response.Write "<a class='admin_btn' href='action.asp?type=delete&tbId=" & lngInfoID & "'>ɾ��</a>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				Response.Write "</tr>" & vbCrLf
				
				i = i + 1
				rsInfo.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsInfo.State = 1 Then rsInfo.Close
		Set rsInfo = Nothing
		
	function getState(i)
		dim strRe
		i = trim(i & "")
		select case i
			case "0"
			strRe = "<span style=""color:red; font-weight:bold;"">&times;</span>"
			
			case "1"
			strRe = "<span style=""color:green; font-weight:bold;"">&radic;</span>"
			
			case else
			strRe = ""
		end select
		getState = strRe
	end function
%>
        </tbody>
      </table>
    </div>
    <div id="PaginationPanel">
<%
	If bPagination Then
		strQuery = inHTML(strQuery)
		Response.Write Pagination(strQuery, lngPageCount, lngCurrPage, lngPageSize)
	End If
%>
    </div>
  </div>
  </div>
</form>
</body>
</html>
<%
	Call CloseConn()
%>
