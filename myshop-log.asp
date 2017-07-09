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
    <div class="container" id="app">
     <div class="panel panel-default">
      <div class="panel-heading">借阅历史记录</div>
      <div class="panel-body" style="padding:0;" v-load-more="load_more">
       <form class="form-myshop" id="form-myshop-log" method="post">
            <table class="table table-striped">
                    <thead>
                        <th>#</th>
                        <th>书名</th>
                        <th>送货时间</th>
                    </thead>
                    <tbody>
                        <tr v-for="(item, index) in list">
                            <td>{{index+1}}</td>
                            <td>{{item.Title}}</td>
                            <td>{{item.returnedDate}}</td>
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
    <script src="https://cdn.bootcss.com/vue/2.3.4/vue.min.js"></script>
    <script src="https://cdn.bootcss.com/vue-resource/1.3.4/vue-resource.min.js"></script>
    <script type="text/javascript" src="<%=gstrInstallDir%>js/vue-load-more.js"></script>
    <script type="text/javascript">
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
                current_page: 1,
                max_page: 1,
                loading: false,
            },
            methods:{
                get_list:function(){
                    var vm = this;
                    vm.$http.get('service/myshop-log.asp?PageNum=' + vm.current_page).then(response => {
                        response.json().then(json => {
                            console.log(json);
                            if(json.state == 0){
                                vm.max_page = json.data.maxpage;
                                if(vm.current_page == 1){
                                    vm.list = json.body;
                                }else{
                                    vm.list = vm.list.concat(json.body);
                                } 
                            }
                            setTimeout(function() {
                                vm.loading = false;
                            }, 1000);
                        });
                    });
                },
                load_more:function(){
                    if(!this.loading){
                            if(this.current_page < this.max_page){
                                this.loading = true;
                                this.current_page++;
                                this.get_list();
                            }
                        }
                },
            },
        });
    </script>
  </body>
</html>