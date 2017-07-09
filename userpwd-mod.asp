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
	dim strUserAcc, lngUserId
	dim rsUser, strsql, lngstate, strMsg
	dim strexpdate_s, strexpdate_e, strUserSex, strUserTel, strUserName, struserAdd, lnguseCounts, lngmaxuseCounts, lngmaxuseCountsTemp
	
	strUserAcc = trim(Session("useracc")&"")
	lngUserId = ConvertLong(Session("UserId")&"")

	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	strsql = "select * from user_t where useracc = '" & strUserAcc & "' and UserId = " & lngUserId
	if rsUser.state = 1 then rs.close
	rsuser.open strsql,conn,1,1
	if not(rsUser.bof or rsUser.eof) then
		strUserName = trim(rsuser("UserName") & "")
		strUserSex = trim(rsuser("UserSex") & "")
		struserAdd = trim(rsuser("userAdd") & "")
		strUserTel = trim(rsuser("UserTel") & "")
		strhobby = trim(rsuser("hobby") & "")
	else
		response.write "出错了"
	end if
%>
  <body style="padding-top:50px;">
	<!--#include file="header.asp"-->
    
    <div class="container">
      <form class="form-usermod" id="form-usermod" method="post">
        <div class="row show-grid">
          <div class="col-xs-6">帐号：<%=strUserAcc%></div>
          <div class="col-xs-6">电话：<%=strUserTel%></div>
        </div>
        <hr>
        <div class="row show-grid">
          <div class="col-xs-12">
            <div class="input-group input-group-lg">
              <span class="input-group-addon" id="sizing-chkuserpwd">密码校验</span>
              <input type="password" class="form-control" placeholder="输入密码进行校验" name="chkuserpwd" id="chkuserpwd" aria-describedby="sizing-chkuserpwd"  required autofocus>
            </div>
          </div>
        </div>
        <hr>
        <div class="row show-grid">
          <div class="col-xs-12">
            <div class="input-group input-group-lg">
              <span class="input-group-addon" id="sizing-newuserpwd">新的密码</span>
              <input type="password" class="form-control" placeholder="输入新密码" name="newuserpwd" id="newuserpwd" aria-describedby="sizing-chkuserpwd"  required autofocus>
            </div>
          </div>
        </div>
        <div class="row show-grid">
          <div class="col-xs-12">
            <div class="input-group input-group-lg">
              <span class="input-group-addon" id="sizing-chknewuserpwd">重复密码</span>
              <input type="password" class="form-control" placeholder="重复输入新密码" name="chknewuserpwd" id="chknewuserpwd" aria-describedby="sizing-chkuserpwd"  required autofocus>
            </div>
          </div>
        </div>
        <hr>
        <div class="row show-grid">
          <div class="col-xs-12 text-right">
              <button type="submit" class="btn btn-info">
                 <span class="glyphicon glyphicon-floppy-save" aria-hidden="true"></span>
                 保存修改
              </button>
          </div>
        </div>

        <hr>
        <div class="btn-group btn-group-justified" role="group" aria-label="...">
          <div class="btn-group" role="group">
            <button type="button" class="btn btn-default" id="userinfo-mod">
                <span class="glyphicon glyphicon-pencil" aria-hidden="true"></span>
                修改资料
            </button>
          </div>
          <div class="btn-group" role="group">
            <button type="button" class="btn btn-default" id="service-tel">
                <span class="glyphicon glyphicon-phone-alt" aria-hidden="true"></span>
                联系客服
            </button>
          </div>
          <div class="btn-group" role="group">
            <button type="button" class="btn btn-default" id="userpwd-mod">
                <span class="glyphicon glyphicon-lock" aria-hidden="true"></span>
                修改密码
            </button>
          </div>
        </div>
      </form>

    </div>
    <!-- /container -->
	<!--#include file="footer.asp"-->
    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://cdn.bootcss.com/jquery/1.12.4/jquery.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="https://cdn.bootcss.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
    <script>
		$("#form-usermod").submit( function(){
			if ($("#newuserpwd").val() == $("#chknewuserpwd").val()){
				$.ajax({
				url:  "service/userpwd-mod.asp", 
				data:$("#form-usermod").serialize(), 
				dataType:'json', 
				type:'post', 
				success:function(data){
					if (data.state == 0) {alert("修改密码成功");location.href = "main.asp";}
					if (data.state == 1) {alert(data.msg);}
					},
				error: function(error) {
					alert("出错了" + error);
					console.log(error);
					return false;
					}
				});
			} 
			else {
				alert("两次密码不一致！");
				$("#newuserpwd").val("");
				$("#chknewuserpwd").val("");
				$("#newuserpwd").focus();
				return false;
			}
		});

		$(function() { 
			$("#nav-userinfo").click(function(){location.href = "main.asp";});
			$("#userinfo-mod").click(function(){location.href = "userinfo-mod.asp";});
			$("#userpwd-mod").click(function(){location.href = "userpwd-mod.asp";});
			$("#service-tel").click(function(){location.href = "tel:<%=gstrServiceTel%>";});
		});
	</script>
  </body>
</html>