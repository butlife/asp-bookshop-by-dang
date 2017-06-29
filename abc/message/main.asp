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
	if(!confirm("您确定要删除本条记录吗?")){
		return;
	}
	location.href = "action.asp?Type=delete&tbId=" + id;
}

function btnView_Click(id) {
	var url = "view.asp?Id=" + id;
	openWindow(url, 500, 450)
}

function btnMsg_Click(id) {
	var url = "msg.asp?messid=" + id + "&s=" + Math.random();
	openDialog(url, 450, 300)
}

</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">留言，投诉，建议列表</div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">序号</th>
            <th nowrap="nowrap">查看</th>
            <th nowrap="nowrap">姓名</th>
            <th nowrap="nowrap">电话</th>
            <!--<th nowrap="nowrap">移动电话</th>
            <th nowrap="nowrap">状态</th>-->
            <th nowrap="nowrap">IP</th>
            <th nowrap="nowrap">时间</th>
            <th nowrap="nowrap">操作</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		'获取数据，显示列表
		Dim rsmess, i, strSql, lngregid, strrecontent
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination, strQuery
		Dim strregname, strsex, strage, strtelephone, strremark, strIP, strMakedate, strtype, strstate
		
		Set rsmess = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From message_t Order By regid desc"
		rsmess.Open strSql, conn, 1, 1
		
		If Not (rsmess.Bof Or rsmess.Eof) Then
			bPagination = True
			'分页
			lngPageSize = glngPageSize
			lngRecordCount = rsmess.RecordCount
			rsmess.PageSize = lngPageSize
			lngPageCount = rsmess.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsmess.AbsolutePage = lngCurrPage
			
			i = 0
			
			'显示列表
			Do While Not (rsmess.Bof Or rsmess.Eof) 
				lngregid = ConvertLong(rsmess("regid") & "")
				strregname = trim(rsmess("regname") & "")
				strtype = trim(rsmess("stype") & "")
				if strtype = "1" then strtypeName = "留言，投诉，建议"
				if strtype = "2" then strtypeName = "会计报名"
				strtelephone = trim(rsmess("telephone") & "")
				strIP = trim(rsmess("IP") & "")
				strMakedate = trim(rsmess("makedate") & "")
				strrecontent = trim(rsmess("recontent") & "")
				strremark = trim(rsmess("remark") & "")
				'strstate = "<span style='color:#0f0;'>己回复</span>"
				'if strrecontent = "" then strstate = "<span style='color:#f00;'>待回复</span>"
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""left""><span class='admin_view' onclick='btnView_Click(" & lngregid & ");'>查看</span></td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strregname) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strtelephone) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strIP) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & strMakedate & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<span class='admin_btn' onclick='btnDelete_Click(" & lngregid & ");'>删除</span>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				Response.Write "</tr>" & vbCrLf
				i = i + 1
				rsmess.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsmess.State = 1 Then rsmess.Close
		Set rsmess = Nothing
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
