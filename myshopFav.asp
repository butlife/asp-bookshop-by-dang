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
    <title>我的收藏夹-<%=gstrKeyWords%></title>

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
	dim lngUserId
	
	'单次最大借书数量
	lnguseCounts = getUserInfo("useCounts")
%>
  <body style="padding-top:50px;">
	<!--#include file="header.asp"-->
    <div class="container">
       <form class="form-myshopFav" id="form-myshopFav" method="post">
     <div class="panel panel-default">
      <div class="panel-heading">收藏夹</div>
      <div class="panel-body" style="padding:0;">
            <table class="table table-striped">
                    <thead>
                        <th>#</th>
                        <th>书名</th>
                        <th>收藏时间</th>
                    </thead>
                    <tbody>
                        <tr>
                            <td><input type="checkbox" name="favId" value="1"></td>
                            <td>国王的新衣<span class="badge">4</span></td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td><input type="checkbox" name="favId" value="2"></td>
                            <td>国王的新衣<span class="badge">4</span></td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td><input type="checkbox" name="favId" value="3"></td>
                            <td>唐诗200首<span class="badge">4</span></td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td><input type="checkbox" name="favId" value="4"></td>
                            <td>国王的新衣<span class="badge">4</span></td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td><input type="checkbox" name="favId" value="3"></td>
                            <td>唐诗200首<span class="badge">4</span></td>
                            <td>2017/06/23</td>
                        </tr>
                        <tr>
                            <td><input type="checkbox" name="favId" value="4"></td>
                            <td>国王的新衣<span class="badge">4</span></td>
                            <td>2017/06/23</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <div class="panel-footer">
                <div class="row show-grid">
                  <div class="col-xs-12 text-right">
                
                      <button type="button" class="btn btn-info" id="unFav-btn">
                         <span class="glyphicon glyphicon-trash" aria-hidden="true"></span>
                         取消收藏
                      </button>
        
                      <button type="button" class="btn btn-info" id="finishOrder-btn">
                         <span class="glyphicon glyphicon-shopping-cart" aria-hidden="true"></span>
                         确认下单
                      </button>
                  </div>
                </div>
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
		$(function() { 
			$("#nav-userinfo").click(function(){location.href = "main.asp";});
			//取消收藏
			$("#unFav-btn").click(function(){
				var ArrinfoId = $("input:checked");
				if (ArrinfoId.length <=0) {
					alert("请选中要准备 【取消收藏】 的记录!");
					return false;
				}
				if(confirm("【取消收藏】 记录将不能恢复,继续吗?")) {
					$.ajax({
					url:  "service/myshopFav-unFav.asp", 
					data:$("#form-myshopFav").serialize(), 
					dataType:'json', 
					type:'post', 
					success:function(data){
						if (data.state == 0) {alert(data.msg);location.href = "myshopFav.asp";}
						if (data.state == 1) {alert(data.msg);}
						},
					error: function(error) {
						alert("出错了" + error);
						console.log(error);
						return false;
						}
					});
				}
			});
			//下订单操作
			$("#finishOrder-btn").click(function(){

				var ArrinfoId = $("input:checked");
				if (ArrinfoId.length<=0) {
					alert("请选中要准备 【确认下单】 的记录!");
					return false;
				}
				if (ArrinfoId.length > <%=lnguseCounts%>) {
					alert("单次最大【确认下单】 为 <%=lnguseCounts%> 本，请重新选择!");
					return false;
				}
				if(confirm("【确认下单】 后，订单不可修改,确认继续吗?")) {
					//alert("确认下单：判断是否有选，是否超过该会员单次最大订书数量，弹出确认下单框，提示下单后不可修改");
					//从收藏夹下订单提交
					$.ajax({
					url:  "service/myshopFav-order.asp", 
					data:$("#form-myshopFav").serialize(), 
					dataType:'json', 
					type:'post', 
					success:function(data){
						if (data.state == 0) {alert(data.msg);location.href = "myshopFav.asp";}
						if (data.state == 1) {alert(data.msg);}
						},
					error: function(error) {
						alert("出错了" + error);
						console.log(error);
						return false;
						}
					});
				}
			});
		});
	</script>
  </body>
</html>