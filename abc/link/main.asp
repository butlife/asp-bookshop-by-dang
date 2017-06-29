<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/Pagination.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!--#include file="../../LyteBox/LyteBox.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>js/common.js" type="text/javascript"></script>

<script language="javascript" type="text/javascript">
function Body_Load(){
}
</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">链接列表</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">链接添加</a>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">&nbsp;</th>
            <th nowrap="nowrap">名称</th>
            <th nowrap="nowrap">网址</th>
            <th nowrap="nowrap">更新时间</th>
            <th nowrap="nowrap">操作</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		'获取数据，显示列表
		Dim lngId, strtitle, strhttpurl, strpicurl, strmakedate
		Dim rsAds, i, strSql, strQuery
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
		
		Set rsAds = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From link_t order by Id "
		rsAds.Open strSql, conn, 1, 1
		
		If Not (rsAds.Bof Or rsAds.Eof) Then
			bPagination = True
			'分页
			lngPageSize = glngPageSize
			lngRecordCount = rsAds.RecordCount
			rsAds.PageSize = lngPageSize
			lngPageCount = rsAds.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsAds.AbsolutePage = lngCurrPage
			
			i = 0
			
			'显示列表
			Do While Not (rsAds.Bof Or rsAds.Eof) 

				lngId = ConvertLong(rsAds("Id") & "")
				strtitle = trim(rsAds("title") & "")
				strhttpurl = trim(rsAds("httpurl") & "")
				strpicurl = trim(rsAds("picurl") & "")
				strmakedate = Format_Time(rsAds("makedate"),1)
				
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""left"">" & TdString(strtitle) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""left""><a href='" & strhttpurl & "' target='_blank'>" & TdString(strhttpurl) & "</a></td>" & vbCrLf
				'Response.Write "<td nowrap=""nowrap"" align=""center""><a href='../../uppic/big/" & strpicurl & "' rel=""lytebox"">" & TdString(strpicurl) & "</a></td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strmakedate) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngId & "'>修改</a>" & vbCrLf
				Response.Write "<a class='admin_btn' href='action.asp?type=delete&tbId=" & lngId & "'>删除</a>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				Response.Write "</tr>" & vbCrLf
				i = i + 1
				rsAds.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsAds.State = 1 Then rsAds.Close
		Set rsAds = Nothing
%>
        </tbody>
      </table>
    </div>
    <div id="PaginationPanel">
<%
		'显示分页效果
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
