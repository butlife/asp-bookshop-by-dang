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

function search() {
	var formMain = document.forms[0];
	formMain.action = "#";
	formMain.submit();
}

function btnManyDelete_Click(){
	var oCltNAMEs = document.getElementsByName('CHK');
	var bPass = false;
	var DelConunt = 0;
	for(var i=0; i<oCltNAMEs.length; i++) {
		var oCltState = oCltNAMEs[i].checked;
		if (oCltState == true) {
			DelConunt= DelConunt+1;
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
		document.forms[0].action="manyaction.asp?type=manyDelete";
		document.forms[0].submit();
	}
}

function btnManySend_Click(){
	var oCltNAMEs = document.getElementsByName('CHK');
	var bPass = false;
	var DelConunt = 0;
	for(var i=0; i<oCltNAMEs.length; i++) {
		var oCltState = oCltNAMEs[i].checked;
		if (oCltState == true) {
			DelConunt= DelConunt+1;
			bPass = true;
		}
		
	}
	if (bPass == false) {
		alert("请选中要准备【发货】的记录!");
		return false;
	}
	//删除的总条数
	//提交到删除页
	if(confirm("确定【发货】吗？此操作不可恢复,继续吗?")) {
		document.forms[0].action="manyaction.asp?type=manySend";
		document.forms[0].submit();
	}
}


</script>
</head>
<body onLoad="Body_Load();">
<form action="#" method="post" name="form1">
    <input type="hidden" name ="main" id="main" value="0" />
  <div id="headPanel">订单列表</div>
  <div id="buttonPanel">
    <input type="button" class="Button" onClick="btnManySend_Click();" value="批量发货" />
    <input type="button" class="Button" onClick="btnManyDelete_Click();" value="批量删除" />
    <span class="red">||</span>
    <a class="" href="main.asp">已下单</a>
    <a class="" href="main1.asp">已发货</a>
    <a class="" href="main2.asp">待收回</a>
    <a class="" href="main3.asp">己完结</a>
    <span class="red">||</span>
<%
	Dim lngInfoID, strtitle, strKeyWords, lngShopID, lnguserId, strUseracc, strusername, strusertel, strAdddate, lngshopState
	Dim rsShop, i, strSql, strQuery
	Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
	dim strUserKeyWords, strBookKeyWords

	strUserKeyWords = trim(request("UserKeyWords") & "")
	strBookKeyWords = trim(request("BookKeyWords") & "")
	strQuery = Trim(Request("Query") & "")
	
%>
	会员名或手机：<input type="text" name="UserKeyWords" id="UserKeyWords" value="<%=strUserKeyWords%>" size="15" />
	书名：<input type="text" name="BookKeyWords" id="BookKeyWords" value="<%=strBookKeyWords%>" size="15" />
	<input type="button" onClick="search();" value="查询">
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" width="98%" cellspacing="0">
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap" width="35">&nbsp;</th>
            <th nowrap="nowrap"><input name="MAINCHK" id="MAINCHK" type="checkbox" value="" onClick="cheageBox('MAINCHK','CHK');"></th>
            <th nowrap="nowrap">姓名</th>
            <th nowrap="nowrap">帐号</th>
            <th nowrap="nowrap">手机</th>
            <th nowrap="nowrap">订单书本</th>
            <th nowrap="nowrap">下单时间</th>
            <th nowrap="nowrap">状态</th>
            <!--<th nowrap="nowrap">操作</th>-->
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		
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
		strSql = "Select * From shop_v " & strQuery & " Order By userId, shopId asc "
		rsShop.Open strSql, conn, 1, 1
		
		If Not (rsShop.Bof Or rsShop.Eof) Then
			bPagination = True
			'分页
			lngPageSize = glngPageSize
			lngRecordCount = rsShop.RecordCount
			rsShop.PageSize = lngPageSize
			lngPageCount = rsShop.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsShop.AbsolutePage = lngCurrPage
			
			i = 0
			
			Do While Not (rsShop.Bof Or rsShop.Eof) 

				lngShopID = ConvertLong(rsShop("ShopID") & "")
				lngInfoID = ConvertLong(rsShop("InfoID") & "")
				lnguserId = ConvertLong(rsShop("userId") & "")
				strtitle = trim(rsShop("title") & "")
				strUseracc = trim(rsShop("Useracc") & "")
				strusername = trim(rsShop("username") & "")
				strusertel = trim(rsShop("usertel") & "")
				strAdddate = Format_Time(rsShop("Adddate"),6)
				lngshopState = ConvertLong(rsShop("shopState") & "")
				
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center""><input name=""CHK"" id=""CHK"" type=""checkbox"" value=""" & lngShopID & """></td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strusername) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strUseracc) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strusertel) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strtitle) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strAdddate) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & getState(lngshopState) & "</td>" & vbCrLf
'				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
'				Response.Write "<a class='admin_btn' href='action.asp?type=Send&tbId=" & lngInfoID & "'>发货</a>" & vbCrLf
'				Response.Write "<a class='admin_btn' href='action.asp?type=delete&tbId=" & lngInfoID & "'>删除</a>" & vbCrLf
'				Response.Write "</td>" & vbCrLf
				Response.Write "</tr>" & vbCrLf
				
				i = i + 1
				rsShop.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsShop.State = 1 Then rsShop.Close
		Set rsShop = Nothing
		
	function getState(i)
		dim strRe
		i = trim(i & "")
		select case i
			case "0"
			strRe = "<span style=""color:red; font-weight:bold;"">下单待发货</span>"
			
			case "1"
			strRe = "<span style=""color:green;"">送书待归还</span>"
			
			case "2"
			strRe = "<span style=""color:blue; font-weight:bold;"">归还待确认</span>"
			
			case "3"
			strRe = "<span style="""">己还书完结</span>"
			
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
