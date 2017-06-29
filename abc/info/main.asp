<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/Pagination.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!--#include file="infosort.asp"-->
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

function search() {
	var formMain = document.forms[0];
	formMain.action = "#";
	formMain.submit();
}
</script>
</head>
<body onLoad="Body_Load();">
<form action="#" method="post" name="form1">
  <div id="headPanel">图书列表</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">图书添加</a>
    <a class="aButton" style="cursor:pointer;" onClick="AllDelete('CHK');">批量删除</a>
<%
	Dim lngInfoID, strtitle, bispassed, dtmakedate, strsortname, stradminname, ihit, strauthor, bistop, iorder, strKeyWords, strpicurl
	Dim rsInfo, i, strSql, strQuery, lngiCount
	Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
	dim lnginfosort_CBS_ID, lnginfosort_FM_ID, lnginfosort_NL_ID, lnginfosort_XL_ID, lnginfosort_ZT_ID

	lnginfosort_CBS_ID = ConvertLong(Request("infosort_CBS_ID") & "")
	lnginfosort_FM_ID = ConvertLong(Request("infosort_FM_ID") & "")
	lnginfosort_NL_ID = ConvertLong(Request("infosort_NL_ID") & "")
	lnginfosort_XL_ID = ConvertLong(Request("infosort_XL_ID") & "")
	lnginfosort_ZT_ID = ConvertLong(Request("infosort_ZT_ID") & "")
	strQuery = Trim(Request("Query") & "")
	
	infosort_NL(lnginfosort_NL_ID)
	response.write "&nbsp;"
	infosort_ZT(lnginfosort_ZT_ID)
	response.write "&nbsp;"
	infosort_XL(lnginfosort_XL_ID)
	response.write "&nbsp;"
	infosort_FM(lnginfosort_FM_ID)
	response.write "&nbsp;"
	infosort_CBS(lnginfosort_CBS_ID)
	response.write "&nbsp;&nbsp;"
%>
	<input type="button" onClick="search();" value="查询">
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" width="98%" cellspacing="0">
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap" width="35">&nbsp;</th>
            <th nowrap="nowrap"><input name="MAINCHK" id="MAINCHK" type="checkbox" value="" onClick="cheageBox('MAINCHK','CHK');"></th>
            <th nowrap="nowrap">标题</th>
            <th nowrap="nowrap">库存</th>
            <th nowrap="nowrap">类别</th>
            <th nowrap="nowrap">审核|置顶</th>
            <th nowrap="nowrap">图片</th>
            <th nowrap="nowrap">更新时间</th>
            <th nowrap="nowrap">操作</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		
'		If Trim(strQuery) = "" Then
			strQuery = " Where 1 = 1 "
				If lnginfosort_CBS_ID <> 0 Then
					strQuery = strQuery & " And infosort_CBS_ID = " & lnginfosort_CBS_ID
				End If
				If lnginfosort_FM_ID <> 0 Then
					strQuery = strQuery & " And infosort_FM_ID = " & lnginfosort_FM_ID
				End If
				If lnginfosort_NL_ID <> 0 Then
					strQuery = strQuery & " And infosort_NL_ID = " & lnginfosort_NL_ID
				End If
				If lnginfosort_XL_ID <> 0 Then
					strQuery = strQuery & " And infosort_XL_ID = " & lnginfosort_XL_ID
				End If
				If lnginfosort_ZT_ID <> 0 Then
					strQuery = strQuery & " And infosort_ZT_ID = " & lnginfosort_ZT_ID
				End If
'		Else
'			strQuery = outHTML(strQuery)
'		End If
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From info_t " & strQuery & " Order By istop desc, ispassed desc, iorder desc, infoid desc "
'		response.write strQuery & "="
'		response.write lnginfosort_CBS_ID & "="
'		response.write strSql
		rsInfo.Open strSql, conn, 1, 1
		
		If Not (rsInfo.Bof Or rsInfo.Eof) Then
			bPagination = True
			'分页
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
				'strsortname = Trim(rsInfo("sortname") & "")
				'stradminname = Trim(rsInfo("adminname") & "")
				ihit = ConvertDouble(rsInfo("hit") & "")
				'strauthor = Trim(rsInfo("author") & "")
				iorder = ConvertDouble(rsInfo("iorder") & "")
				'strKeyWords = Trim(rsInfo("KeyWords") & "")
				lngiCount = ConvertLong(rsInfo("iCount") & "")
				strsortname = ""
				strsortname = strsortname & Trim(getSortName("infosort_NL",rsInfo("infosort_NL_id"))) & "|"
				strsortname = strsortname & Trim(getSortName("infosort_ZT",rsInfo("infosort_ZT_id"))) & "|"
				strsortname = strsortname & Trim(getSortName("infosort_XL",rsInfo("infosort_XL_id"))) & "|"
				strsortname = strsortname & Trim(getSortName("infosort_FM",rsInfo("infosort_FM_id"))) & "|"
				strsortname = strsortname & Trim(getSortName("infosort_CBS",rsInfo("infosort_CBS_id")))
				
				strpicurl = Trim(rsInfo("picurl") & "")
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center""><input name=""CHK"" id=""CHK"" type=""checkbox"" value=""" & lngInfoId & """></td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""left"">" & TdString(strtitle) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(lngiCount) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strsortname) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & getState(bispassed) & "|" & getState(bistop) & "</td>" & vbCrLf
				if strpicurl = "" then
					Response.Write "<td nowrap=""nowrap"" align=""center"">&nbsp;</td>" & vbCrLf
				else
					Response.Write "<td nowrap=""nowrap"" align=""center""><a href=""../../uppic/big/" & strpicurl & """ target=""_blank"">图片</a></td>" & vbCrLf
				end if
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(dtmakedate) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngInfoID & "'>修改</a>" & vbCrLf
				If (bistop = 1) Then
					Response.Write "<a class='admin_btn_b' href='action.asp?type=top&tbId=" & lngInfoID & "'>解除</a>" & vbCrLf
				Else
					Response.Write "<a class='admin_btn' href='action.asp?type=top&tbId=" & lngInfoID & "'>置顶</a>" & vbCrLf
				End If
				Response.Write "<a class='admin_btn' href='action.asp?type=delete&tbId=" & lngInfoID & "'>删除</a>" & vbCrLf
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
	function getSortName(tab, id)
		dim rsSort, strsortSql, strRe
		strRe = "&nbsp;"
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		id = ConvertLong(id & "")
		strsortSql = "Select sortname From " & tab & " where sortid =  " & id
		rsSort.Open strsortSql, conn, 1, 1
		if not(rsSort.bof or rsSort.eof) then
			strRe = trim(rsSort("sortname")&"")
		end if
		If rsSort.State = 1 Then rsSort.Close
		Set rsSort = Nothing
		getSortName = strRe
	end function
	
	Call CloseConn()
%>
