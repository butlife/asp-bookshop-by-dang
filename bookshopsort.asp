<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- 上述3个meta标签*必须*放在最前面，任何其他内容都*必须*跟随其后！ -->
	<!--#include file="common/conn-utf.asp"-->
    <!--#include file="common/Function-utf.asp"-->
    <!--#include file="common/safe.asp"-->
    <title>网上书城-类别-<%=gstrKeyWords%></title>

    <!-- Bootstrap -->
    <link rel="stylesheet" href="https://cdn.bootcss.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">
    <link href="<%=gstrInstallDir%>bootstrap337/css/bootstrap.min.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://cdn.bootcss.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
    <style type="text/css">
		.padLeft36 {
			padding-left:36px;
		}
	</style>
  </head>
<%
	function getSortList(Tab, SortName)
		dim rsSort, strsql
		dim strSortName, lngSortId
	
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strsql = "select * from " & Tab & " order by iOrder asc"
		if rsSort.state = 1 then rs.close
		rsSort.open strsql,conn,1,1
    	response.write "<div class=""panel panel-info"">"
            response.write "<div class=""panel-heading"">" & SortName & "</div>"
            response.write "<div class=""list-group"">"
		if not(rsSort.bof or rsSort.eof) then
			response.write "<div class=""list-group"">"
				do while not(rsSort.bof or rsSort.eof)
					lngSortId = ConvertLong(rsSort("SortId") & "")
					strSortName = trim(rsSort("SortName") & "")
					response.write "<a href=""bookshop.asp?sortid=" & lngSortId & """ class=""list-group-item padLeft36"">" & strSortName & "</a>"
					rsSort.movenext
				loop
			response.write "</div>"
		end if
			response.write "</div>"
		response.write "</div>"
	end function
%>
  <body style="padding-top:50px;">
	<!--#include file="header.asp"-->
    <div class="container">
<%
		call getSortList("infosort_NL", "按年龄分类")
		call getSortList("infosort_XL", "按系列分类")
		call getSortList("infosort_ZT", "按主题分类")
		call getSortList("infosort_ZT", "父母专区")
		call getSortList("infosort_CBS", "按出版社分类")
%>
    </div>
    <!-- /container -->
	<!--#include file="footer.asp"-->
    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://cdn.bootcss.com/jquery/1.12.4/jquery.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="https://cdn.bootcss.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
    <script>
		$(function() { 
			$("#nav-userinfo").click(function(){location.href = "main.asp";});
			$("#userinfo-mod").click(function(){location.href = "userinfo-mod.asp";});
			$("#userpwd-mod").click(function(){location.href = "userpwd-mod.asp";});
			$("#service-tel").click(function(){location.href = "tel:<%=gstrServiceTel%>";});
		});
	</script>
  </body>
</html>