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
		if(!confirm("删除会员同时会删除该会员所有借书记录，\n删除的记录将不可恢复,继续吗?")){
			return;
		}
		location.href = "action.asp?Type=delete&tbId=" + id;
	}
</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">会员列表</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">会员添加</a>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">&nbsp;</th>
            <th nowrap="nowrap">姓名</th>
            <th nowrap="nowrap">性别</th>
            <th nowrap="nowrap">帐号</th>
            <th nowrap="nowrap">电话</th>
            <th nowrap="nowrap">有效期</th>
            <th nowrap="nowrap" title="单次借书数量/最大借书次数/会员状态">次数/状态</th>
            <th nowrap="nowrap">操作</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		'获取数据，显示列表
		Dim lngUserId, strUserName, strusersex, struseracc, strexpdate_s, strexpdate_e, lnguseCounts, lngmaxuseCounts, bispassed, strusertel
		Dim rsUser, i, strSql, strQuery
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
		
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From user_t Order By loginDate desc "
		rsUser.Open strSql, conn, 1, 1
		
		If Not (rsUser.Bof Or rsUser.Eof) Then
			bPagination = True
			'分页
			lngPageSize = glngPageSize
			lngRecordCount = rsUser.RecordCount
			rsUser.PageSize = lngPageSize
			lngPageCount = rsUser.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsUser.AbsolutePage = lngCurrPage
			
			i = 0
			
			'显示列表
			Do While Not (rsUser.Bof Or rsUser.Eof) 

				lngUserId = ConvertLong(rsUser("UserId") & "")
				strUserName = trim(rsUser("UserName") & "")
				strusersex = trim(rsUser("usersex") & "")
				struseracc = trim(rsUser("useracc") & "")
				strexpdate_s = Format_Time(rsUser("expdate_s"), 5)
				strexpdate_e = Format_Time(rsUser("expdate_e"), 5)
				lnguseCounts = ConvertLong(rsUser("useCounts") & "")
				lngmaxuseCounts = ConvertLong(rsUser("maxuseCounts") & "")
				bispassed = ConvertLong(rsUser("ispassed") & "")
				strusertel = trim(rsUser("usertel") & "")
				
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strUserName) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strusersex) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(struseracc) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strusertel) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strexpdate_s & "-" & strexpdate_e) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(lnguseCounts & "|" & lngmaxuseCounts & "/" & getState(bispassed)) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngUserId & "'>修改</a>" & vbCrLf
				Response.Write "<a class='admin_btn' href='javascript:btnDelete_Click(" & lngUserId & ")'>删除</a>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				Response.Write "</tr>" & vbCrLf
				i = i + 1
				rsUser.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsUser.State = 1 Then rsUser.Close
		Set rsUser = Nothing
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

	Call CloseConn()
%>
