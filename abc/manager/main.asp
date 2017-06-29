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
</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">管理员列表</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">管理员添加</a>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">序号</th>
            <th nowrap="nowrap">名称</th>
            <th nowrap="nowrap">状态</th>
            <th nowrap="nowrap">添加时间</th>
            <th nowrap="nowrap">上次登陆时间</th>
            <th nowrap="nowrap">上次登陆IP</th>
            <th nowrap="nowrap">登陆次数</th>
            <th nowrap="nowrap">操作</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
          <%
		'获取数据，显示列表
		Dim lngAdminId, strAdminName, bispassed, dtAddTime, strLastLoginIp, dtlogindate, lngLoginCount, strState

		Dim rsAdmin, i, strSql, strQuery
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
		
		Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From manager_t Order By id "
		rsAdmin.Open strSql, conn, 1, 1
		
		If Not (rsAdmin.Bof Or rsAdmin.Eof) Then
			bPagination = True
			'分页
			lngPageSize = glngPageSize
			lngRecordCount = rsAdmin.RecordCount
			rsAdmin.PageSize = lngPageSize
			lngPageCount = rsAdmin.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsAdmin.AbsolutePage = lngCurrPage
			
			i = 0
			
			'显示列表
			Do While Not (rsAdmin.Bof Or rsAdmin.Eof) 

				lngAdminId = ConvertLong(rsAdmin("ID") & "")
				strAdminName = ReplaceBadChar(rsAdmin("AdminName") & "")
				bispassed = ConvertDouble(rsAdmin("ispassed") & "")
				dtAddTime = Trim(rsAdmin("adddate") & "")
				strLastLoginIp = Trim(rsAdmin("sip") & "")
				dtlogindate = Trim(rsAdmin("logindate") & "")
				lngLoginCount = ConvertDouble(rsAdmin("logocount") & "")
				
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If

				
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strAdminName) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & getState(bispassed) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdDate(dtAddTime) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(dtlogindate) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strLastLoginIp) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdNumeric(lngLoginCount) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngAdminId & "'>修改</a>" & vbCrLf
				If (bispassed = 1) Then
					Response.Write "<a class='admin_btn' href='action.asp?type=passed&tbId=" & lngAdminId & "'>弃核</a>" & vbCrLf
				Else
					Response.Write "<a class='admin_btn_b' href='action.asp?type=passed&tbId=" & lngAdminId & "'>审核</a>" & vbCrLf
				End If
				Response.Write "<a class='admin_btn' href='action.asp?type=delete&tbId=" & lngAdminId & "'>删除</a>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				
				Response.Write "</tr>" & vbCrLf
				
				i = i + 1
				rsAdmin.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsAdmin.State = 1 Then rsAdmin.Close
		Set rsAdmin = Nothing
		
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
