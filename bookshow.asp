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
    <title>修改会员信息-<%=gstrKeyWords%></title>

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
	dim rsinfo, strsql, lngstate, strMsg
	dim lnginfoId
	
	lnginfoId = ConvertLong(request("infoId")&"")

	Set rsinfo = Server.CreateObject("ADODB.RecordSet")
	strsql = "select title, content from info_t where infoId = " & lnginfoId
	if rsinfo.state = 1 then rs.close
	rsinfo.open strsql,conn,1,1
	if not(rsinfo.bof or rsinfo.eof) then
		strTitle = trim(rsinfo("Title") & "")
		strContent = trim(rsinfo("Content") & "")
	else
		strTitle = "出错了，请返回刷新后重试。"
		strContent = "出错了，请返回刷新后重试。"
	end if
%>
  <body style="padding-top:50px;">
	<!--#include file="header.asp"-->
    <div class="container">
      <form class="form-usermod" id="form-usermod" method="post">
         <div class="panel panel-default">
          <div class="panel-heading"><%=strTitle%></div>
          <div class="panel-body" style="">
          	<%=strContent%>
          </div>
          <div class="panel-footer" style="text-align:right;">
            <button type="button" class="btn btn-default" id="returnList">
                <span class="glyphicon glyphicon-chevron-left" aria-hidden="true"></span>
                返回列表
            </button>
          </div>
        </div>          
      </form>
    </div>
    <!-- /container -->
    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://cdn.bootcss.com/jquery/1.12.4/jquery.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="https://cdn.bootcss.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
    <script>
		$(function() { 
			$("#nav-userinfo").click(function(){location.href = "main.asp";});
			$("#returnList").click(function(){history.go(-1);});
		});
	</script>
  </body>
</html>