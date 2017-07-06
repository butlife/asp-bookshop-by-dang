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
    <title>正在借阅-<%=gstrKeyWords%></title>

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
    <div class="container" id="app">
       <form class="form-myshop" id="form-myshop" method="post">
      <div class="panel panel-default">
      <div class="panel-heading">正在借阅</div>
      <div class="panel-body" style="padding:0;">
           <table class="table table-striped">
                    <thead>
                        <th>#</th>
                        <th>书名</th>
                        <th>状态</th>
                        <th>下单时间</th>
                    </thead>
                    <tbody>
                        <tr v-for="(item, index) in list">
                            <td>
                                <input type="checkbox" v-show="item.edit" name="shopid" v-bind:value="item.ShopId">
                                <span v-show="!item.edit">#</span>
                            </td>
                            <td>{{item.Title}}</td>
                            <td>{{item.shopStateName}}</td>
                            <td>{{item.AddDate}}</td>
                        </tr>
                    </tbody>
                </table>
          </div>
          <div class="panel-footer">  
                <div class="row show-grid">
                  <div class="col-xs-12 text-right">
                      <button type="button" class="btn btn-info" id="returned-btn">
                         <span class="glyphicon glyphicon-share" aria-hidden="true"></span>
                         申请归还
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
			$("#returned-btn").click(function(){
				var ArrinfoId = $("input:checked");
				if (ArrinfoId.length <=0) {
					alert("请选中要准备 【申请归还】 的记录!");
					return false;
				}
				if(confirm("【申请归还】 操作后将不能恢复,继续吗?")) {
					$.ajax({
					url:  "service/myshop-returned.asp", 
					data:$("#form-myshop").serialize(), 
					dataType:'json', 
					type:'post', 
					success:function(data){
						if (data.state == 0) {alert(data.msg);location.href = "myshop.asp";}
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
     <script src="https://cdn.bootcss.com/vue/2.3.4/vue.min.js"></script>
    <script src="https://cdn.bootcss.com/vue-resource/1.3.4/vue-resource.min.js"></script>
<script>
    Vue.use(VueResource);
        Vue.http.options.emulateJSON = true;

        var app = new Vue({
            el:'#app',
            created:function(){
                console.log('app created');
                this.get_list();
            },
            data:{
                list:[],
            },
            methods:{
                get_list:function(){
                    var vm = this;
                    vm.$http.get('service/myshop.asp').then(response => {
                        response.json().then(json => {
                            console.log(json);
                            if(json.state == 0){
                                vm.list = json.body;
                            }
                        });
                    });
                },
            },
        });
</script>
  </body>
</html>