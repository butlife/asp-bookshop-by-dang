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
<script language="javascript" type="text/javascript">
	function Body_Load(){
	}
	
	function btnDelete_Click(id){
		if(!confirm("删除类别同时会删除该类别下所属图书，\n删除的记录将不可恢复,继续吗?")){
			return;
		}
		location.href = "action.asp?Type=delete&tbId=" + id;
	}
</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">主题类别列表</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">主题类别添加</a>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">&nbsp;</th>
            <th nowrap="nowrap">名称</th>
            <th nowrap="nowrap">排序号</th>
            <th nowrap="nowrap">操作</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		'获取数据，显示列表
		Dim lngSortId, strSortName, lngOrder
		Dim rsSort, i, strSql, strQuery
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
		
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_ZT Order By iorder, sortid desc "
		rsSort.Open strSql, conn, 1, 1
		
		If Not (rsSort.Bof Or rsSort.Eof) Then
			bPagination = True
			'分页
			lngPageSize = glngPageSize
			lngRecordCount = rsSort.RecordCount
			rsSort.PageSize = lngPageSize
			lngPageCount = rsSort.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsSort.AbsolutePage = lngCurrPage
			
			i = 0
			
			'显示列表
			Do While Not (rsSort.Bof Or rsSort.Eof) 

				lngSortId = ConvertLong(rsSort("SortId") & "")
				strSortName = trim(rsSort("SortName") & "")
				lngOrder = ConvertLong(rsSort("iOrder") & "")
				
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center""><a href='../info/main.asp?infosort_ZT_ID=" & lngSortId & "'>" & TdString(strSortName) & "</a></td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdNumeric(lngOrder) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngSortId & "'>修改</a>" & vbCrLf
				Response.Write "<a class='admin_btn' href='javascript:btnDelete_Click(" & lngSortId & ")'>删除</a>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				Response.Write "</tr>" & vbCrLf
				i = i + 1
				rsSort.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsSort.State = 1 Then rsSort.Close
		Set rsSort = Nothing
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
