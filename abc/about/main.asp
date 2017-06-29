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
	if(!confirm("删除的记录将不可恢复,继续吗?")){
		return;
	}
	location.href = "action.asp?Type=delete&tbId=" + id;
}

function btnModify_Click(id) {
	var url = "modify.asp?Id=" + id;
	location.href = url;
}

function manageradd() {
	var url = "add.asp";
	location.href = url;
}

</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">站点信息列表</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">站点信息添加</a>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">序号</th>
            <th nowrap="nowrap">标题</th>
            <!--<th nowrap="nowrap">图片</th>-->
            <th nowrap="nowrap">管理员</th>
            <th nowrap="nowrap">更新时间</th>
            <th nowrap="nowrap">操作</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		Dim rsInfo, i, strSql, strQuery
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
		Dim lngInfoId, strTitle, strAdminName, lngAdminId, dtUpdateTime, strContent, strpicurl
		
		Set rsInfo = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From about_v Order By id "
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

				lngInfoId = ConvertLong(rsInfo("id") & "")
				lngAdminId = ConvertLong(rsInfo("adminid") & "")
				strTitle = trim(rsInfo("title") & "")
				strAdminName = trim(rsInfo("adminname") & "")
				dtUpdateTime = trim(rsInfo("makedate") & "")
				strpicurl = trim(rsInfo("picurl") & "")
				
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
'				Response.Write "<img src=""" & gstrInstallDir & "Images/view.gif"" boder=""0"" style=""cursor:hand;"" onclick=""btnView_Click(" & lngInfoId & ");"" align=""absmiddle"" title=""详细信息"" />" & vbCrLf
'				Response.Write "<img src=""" & gstrInstallDir & "Images/modify.gif"" boder=""0"" style=""cursor:hand;"" onclick=""btnModify_Click(" & lngInfoId & ");"" align=""absmiddle"" title=""修改"" />" & vbCrLf
'				Response.Write "<img src=""" & gstrInstallDir & "Images/delete.gif"" boder=""0"" style=""cursor:hand;"" onclick=""btnDelete_Click(" & lngInfoId & ");"" align=""absmiddle"" title=""删除"" />" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center""><span class='admin_view'>" & TdString(strTitle) & "</span></td>" & vbCrLf
				'Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strpicurl) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strAdminName) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(dtUpdateTime) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngInfoID & "'>修改</a>" & vbCrLf
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
