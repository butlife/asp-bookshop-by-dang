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
    <title>借书记录-<%=gstrKeyWords%></title>

    <!-- Bootstrap -->
    <link rel="stylesheet" href="https://cdn.bootcss.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">
    <link href="<%=gstrInstallDir%>bootstrap337/css/bootstrap.min.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://cdn.bootcss.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>
<%
	dim strUserAcc, lngUserId
	dim rsUser, strsql, lngstate, strMsg
	dim strexpdate_s, strexpdate_e, strUserSex, strUserTel, strUserName, struserAdd, lnguseCounts, lngmaxuseCounts, lngmaxuseCountsTemp
	
	strUserAcc = trim(Session("useracc")&"")
	lngUserId = ConvertLong(Session("UserId")&"")

	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	strsql = "select * from user_t where useracc = '" & strUserAcc & "' and UserId = " & lngUserId
	if rsUser.state = 1 then rs.close
	rsuser.open strsql,conn,1,3
	if not(rsUser.bof or rsUser.eof) then
		strUserName = trim(rsuser("UserName") & "")
		strUserSex = trim(rsuser("UserSex") & "")
		struserAdd = trim(rsuser("userAdd") & "")
		strexpdate_s = Format_Time(rsuser("expdate_s"),4)
		strexpdate_e = Format_Time(rsuser("expdate_e"),4)
		strUserTel = trim(rsuser("UserTel") & "")
		lnguseCounts = ConvertLong(rsuser("useCounts") & "")
		lngmaxuseCounts = ConvertLong(rsuser("maxuseCounts") & "")
		lngmaxuseCountsTemp = ConvertLong(rsuser("maxuseCountsTemp") & "")
	else
		response.write "出错了"
	end if
%>
  <body style="padding-top:50px;">
	<!--#include file="header.asp"-->
    <div class="container">
    
    
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
			$("#nav-shopsort").click(function(){location.href = "bookshopsort.asp";});
		});
	</script>
  </body>
</html>