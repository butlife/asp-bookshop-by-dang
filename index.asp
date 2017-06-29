<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- 上述3个meta标签*必须*放在最前面，任何其他内容都*必须*跟随其后！ -->
	<!--#include file="common/conn-utf.asp"-->
    <!--#include file="common/Function-utf.asp"-->
    <title><%=gstrKeyWords%>-会员登录</title>

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
  <body>

    <div class="container">

      <form class="form-signin" id="form-signin" method="post">
        <h2 class="form-signin-heading">会员登录</h2>
        <hr>
        <div class="input-group input-group-lg">
          <span class="input-group-addon" id="sizing-userAcc"><span class="glyphicon glyphicon-user" aria-hidden="true"></span></span>
          <input type="text" class="form-control" placeholder="用户帐号" name="userAcc" id="userAcc" aria-describedby="sizing-addon1"  required autofocus>
        </div>
        
        <div class="input-group input-group-lg">
          <span class="input-group-addon" id="sizing-userpwd"><span class="glyphicon glyphicon-lock" aria-hidden="true"></span></span>
          <input type="password" class="form-control" placeholder="用户密码" name="userpwd" id="userpwd" required>
        </div>
        <hr>
        <div class="btn-group btn-group-justified" role="group" aria-label="...">
          <div class="btn-group" role="group">
            <button type="submit" class="btn btn-primary">登录</button>
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
		$("#form-signin").submit( function(){
			$.ajax({
			url:  "service/login.asp", 
			data:$("#form-signin").serialize(), 
			dataType:'json', 
			type:'post', 
			success:function(data){
				if (data.state == 0) {location.href = "main.asp";}
				if (data.state == 1) {alert(data.msg);}
				},
			error: function(error) {
				alert("出错了" + error);
				console.log(error);
				}
			});
		});
	});
	</script>
  </body>
</html>