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
    <title>借阅历史记录-<%=gstrKeyWords%></title>

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
  <body style="padding-top:50px;">
	<!--#include file="header.asp"-->
    <div class="container">
     <div class="panel panel-default">
      <div class="panel-heading">借阅历史记录</div>
      <div class="panel-body" style="padding:0;">
       <form class="form-myshop" id="form-myshop-log" method="post">
            <table class="table table-striped">
                    <thead>
                        <th>#</th>
                        <th>书名</th>
                        <th>状态</th>
                        <th>送货时间</th>
                    </thead>
                    <tbody>
                        <tr>
                            <td>1</td>
                            <td>国王的新衣</td>
                            <td>己归还</td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td>2</td>
                            <td>国王的新衣</td>
                            <td>己送货</td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td>3</td>
                            <td>唐诗200首</td>
                            <td>己下单</td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td>4</td>
                            <td>国王的新衣</td>
                            <td>己完结</td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td>5</td>
                            <td>唐诗200首</td>
                            <td>己送货</td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td>6</td>
                            <td>国王的新衣</td>
                            <td>己送货</td>
                            <td>2017/06/23</td>
                        </tr>
                    </tbody>
                </table>
          </form>
      </div>
    </div>   
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
		});
	</script>
  </body>
</html>