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
<style>
	.tableContainer table {
		border-collapse:collapse;
	}
	.tableContainer td{
		text-align:left;
		height:25px;
		text-indent:5px;
	}
	.tableContainer th {
		height:25px;
		background:#dcdcdc;
		text-align:left;
		text-indent:5px;
	}
</style>
<script language="javascript" type="text/javascript">
function Body_Load(){
}

function Frmsearch() {
	var formMain = document.forms[0];
	formMain.action = "#";
	formMain.submit();
}

function btnManyDelete_Click(frm){
	var tempFrm = document.getElementsByName(frm)[0];
	var bPass = false;
	var tempChk = tempFrm.shopId;
	for(var i=0; i<tempChk.length; i++) {
		var oCltState = tempChk[i].checked;
		if (oCltState == true) {
			bPass = true;
		}
	}
	if (bPass == false) {
		alert("请选中要准备 【删除】 的记录!");
		return false;
	}
	//删除的总条数
	//提交到删除页
	if(confirm("【删除】 记录将不能恢复,继续吗?")) {
		tempFrm.action="manyaction.asp?type=manyDelete";
		tempFrm.submit();
	}
}

function allchk(frm){
	var tempFrm = document.getElementsByName(frm)[0];
	var tempMainChk = tempFrm.mainchk;
	var tempChk = tempFrm.shopId;
	for(var i=0; i<tempChk.length; i++) {
		tempChk[i].checked = tempMainChk.checked;
	}
}


</script>
</head>
<body onLoad="Body_Load();">
  <div id="headPanel">全部订单-订单列表</div>
  <div id="buttonPanel">
<form action="#" method="post" name="form1">
    <input type="hidden" name ="main" id="main" value="all" />
    <input type="hidden" name="Query" value="">
    <input type="hidden" name="Page" value="">
    <a class="" href="main.asp">已下单</a>
    <a class="" href="main1.asp">已发货</a>
    <a class="" href="main2.asp">待收回</a>
    <a class="" href="main3.asp">己完结</a>
    <a class="" href="mainAll.asp">全部</a>
    <span class="red">||</span>
<%
	Dim lngInfoID, strtitle, strKeyWords, lngShopID, lnguserId, strUseracc, strusername, strusertel, strAdddate, lngshopState
	Dim rsShop, i, strSql, strQuery, strArruserId
	Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
	dim strUserKeyWords, strBookKeyWords

	strUserKeyWords = trim(request("UserKeyWords") & "")
	strBookKeyWords = trim(request("BookKeyWords") & "")
	strQuery = Trim(Request("Query") & "")
	
	strArruserId = "0"
	
'		If Trim(strQuery) = "" Then
		strQuery = " Where 1 = 1 "
		If strUserKeyWords <> "" Then
			strQuery = strQuery & " And (username like '%" & strUserKeyWords & "%' or useracc like '%" & strUserKeyWords & "%' or userTel like '%" & strUserKeyWords & "%')"
		End If
		
		If strBookKeyWords <> "" Then
			strQuery = strQuery & " And (title like '%" & strBookKeyWords & "%' or content like '%" & strBookKeyWords & "%')"
		End If
'		Else
'			strQuery = outHTML(strQuery)
'		End If
	Set rsShop = Server.CreateObject("ADODB.RecordSet")
	strSql = "Select userId From shop_v " & strQuery & " Order By shopId asc "
	rsShop.Open strSql, conn, 1, 1
	If Not (rsShop.Bof Or rsShop.Eof) Then
		Do While Not (rsShop.Bof Or rsShop.Eof) 
			strArruserId = strArruserId & ", " & rsshop("userid")
			rsShop.MoveNext
		Loop
	end if

	if rsShop.state = 1 then rsShop.close
	set rsShop = nothing
%>
	会员名或手机：<input type="text" name="UserKeyWords" id="UserKeyWords" value="<%=strUserKeyWords%>" size="15" />
	书名：<input type="text" name="BookKeyWords" id="BookKeyWords" value="<%=strBookKeyWords%>" size="15" />
	<input type="button" onClick="Frmsearch();" value="查询">
</form>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">

<%
	dim rsUser, strUserSql
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	strUserSql = "Select userId, username, useracc, usertel, expdate_s, expdate_e, useCounts, maxuseCounts, maxuseCountsTemp, useradd From user_t where userid in (" & strArruserId & ") "
	dim username, useracc, usertel, expdate_s, expdate_e, useCounts, maxuseCounts, maxuseCountsTemp, useradd	
	if rsUser.state = 1 then rsUser.close
	rsUser.open strUserSql, conn,1,1
		If Not (rsUser.Bof Or rsUser.Eof) Then
			bPagination = True
			'分页
			'lngPageSize = glngPageSize
			lngPageSize = 3
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
			
			Do While Not (rsUser.Bof Or rsUser.Eof) 

				lnguserId = ConvertLong(rsUser("userId") & "")
				username = trim(rsUser("username") & "")
				useracc = trim(rsUser("useracc") & "")
				usertel = trim(rsUser("usertel") & "")
				expdate_s = Format_Time(rsUser("expdate_s"),4)
				expdate_e = Format_Time(rsUser("expdate_e"),4)
				useCounts = ConvertLong(rsUser("useCounts") & "")
				maxuseCounts = ConvertLong(rsUser("maxuseCounts") & "")
				maxuseCountsTemp = ConvertLong(rsUser("maxuseCountsTemp") & "")
				useradd = trim(rsUser("useradd") & "")
%>
<form method="post" name="userOrderFrm<%=lnguserId%>" action="">
<input type="hidden" name="userid" id="userid" value="<%=lnguserId%>" />
<input type="hidden" name ="main" id="main" value="all" />
<table width="96%" border="1" cellspacing="0" cellpadding="0" <%If i Mod 2 = 0 Then response.write "align=""left""" else response.write "align=""right""" end if%>>
  <tr>
    <th>姓名</th>
    <th>帐号</th>
    <th>电话</th>
    <th>有效期</th>
  </tr>
  <tr>
    <td><%=username%></td>
    <td><%=useracc%></td>
    <td><%=usertel%></td>
    <td><%=expdate_s%> - <%=expdate_e%></td>
  </tr>
  <tr>
    <th>次数</th>
    <th colspan="3">地址</th>
    </tr>
  <tr>
    <td>剩<%=maxuseCountsTemp%>/总<%=maxuseCounts%></td>
    <td colspan="3"><%=useradd%></td>
    </tr>
  <tr>
    <td colspan="4"><%=getShopList(lnguserId,99)%></td>
    </tr>
</table>
</form>    
<%				
				i = i + 1
				rsUser.MoveNext
				'If i >= lngPageSize Then Exit Do
				If i >= 3 Then Exit Do
			Loop
		End If
		
		If rsUser.State = 1 Then rsUser.Close
		Set rsUser = Nothing
		
%>
        </tbody>
      </table>
    </div>
    <div id="PaginationPanel">
<form action="#" method="post">
<%
	If bPagination Then
		strQuery = inHTML(strQuery)
		Response.Write Pagination(strQuery, lngPageCount, lngCurrPage, lngPageSize)
	End If
%>
</form>
    </div>
  </div>
  </div>

</body>
</html>
<%
	function getShopList(uid, shopstate)
		uid = ConvertLong(uid & "")
		shopstate = ConvertLong(shopstate & "")
		dim rsShopList, strshopsql, i
		dim lngshopId, strTitle, strAdddate, lngshopstate, strState, strReturnedDate, strSendDate, strfinishdate
		Set rsShopList = Server.CreateObject("ADODB.RecordSet")
		if shopstate = 99 then
			strshopsql = "Select shopid, title, adddate, senddate, returneddate, finishdate, shopstate From shop_v where userid = " & uid
		else
			strshopsql = "Select shopid, title, adddate, senddate, returneddate, finishdate, shopstate From shop_v where shopstate = " & shopstate & " and userid = " & uid
		end if
		if rsShopList.state = 1 then rsShopList.close
		rsShopList.open strshopsql, conn, 1,1
		response.write "<table width=""100%"" border=""1"" cellspacing=""0"" cellpadding=""0"">"
		response.write "<tr>"
		response.write "<th><input type=""checkbox"" name=""mainchk"" onClick=""allchk('userOrderFrm" & uid & "');"" /></th>"
		response.write "<th>书名</th>"
		response.write "<th>时间</th>"
		response.write "<th>状态</th>"
		response.write "<th style=""width:80px;"">操作</th>"
		response.write "</tr>"
		i=0
		do while not(rsShopList.bof or rsShopList.eof)
			lngshopId = ConvertLong(rsShopList("shopId") & "")
			strTitle = trim(rsShopList("Title") & "")
			strAdddate = Format_Time(rsShopList("adddate"),6)
			strSendDate = Format_Time(rsShopList("SendDate"),6)
			strReturnedDate = Format_Time(rsShopList("returneddate"),6)
			strfinishdate = Format_Time(rsShopList("finishdate"),6)
			lngshopstate = ConvertLong(rsShopList("shopstate") & "")
			strState = getState(lngshopstate)
			response.write "<tr>"
			response.write "<td><input type=""checkbox"" name=""shopId"" value=""" & lngshopId & """ /></td>"
			response.write "<td>" & strTitle & "</td>"
			response.write "<td>" & strAdddate & " <span class=""red""> &gt; </span> " & strSendDate & " <span class=""red""> &gt; </span>" & strReturnedDate & " <span class=""red""> &gt; </span>" & strfinishdate & "</td>"
			response.write "<td>" & strState & "</td>"
			if i = 0 then
				response.write "<td rowspan=""999"" style=""text-align:center;"">"
				
				response.write "<input type=""button"" class=""Button"" onClick=""btnManyDelete_Click('userOrderFrm" & uid & "');"" value=""删除"" />"
				
				response.write "</td>"
			end if
			response.write "</tr>"
			i=i+1
			rsShopList.movenext
		loop
		response.write "</table>"
		if rsShopList.state = 1 then rsShopList.close
		set rsShopList = nothing
	end function

	function getState(i)
		dim strRe
		i = trim(i & "")
		select case i
			case "0"
			strRe = "<span style=""color:red; font-weight:bold;"">己下单待发货</span>"
			
			case "1"
			strRe = "<span style=""color:green;"">己送书待归还</span>"
			
			case "2"
			strRe = "<span style=""color:blue; font-weight:bold;"">己归还待确认</span>"
			
			case "3"
			strRe = "<span style="""">己还书订单完结</span>"
			
			case else
			strRe = ""
		end select
		getState = strRe
	end function
	
	Call CloseConn()
%>
