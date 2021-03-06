<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- 上述3个meta标签*必须*放在最前面，任何其他内容都*必须*跟随其后！ -->
	<!--#include file="common/conn-utf.asp"-->
    <!--#include file="common/Function-utf.asp"-->
    <title>网上书城-书库-<%=gstrKeyWords%></title>

    <!-- Bootstrap -->
    <link rel="stylesheet" href="https://cdn.bootcss.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">
    <link href="<%=gstrInstallDir%>bootstrap337/css/bootstrap.min.css" rel="stylesheet">
    <link href="<%=gstrInstallDir%>bootstrap337/dropload/dropload.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://cdn.bootcss.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
	<![endif]-->
        <style>
            .shopbook-img {
                width:150px;
                height:150px;
            }
        </style>
    </head>
    <%
        dim strKeyWords, lngSortId
        strKeyWords = trim(request("bookKeyWords") & "")
        lngSortId = ConvertLong(request("SortId") & "")
        
    %>
    <body style="padding-top:50px;">
        <!--#include file="header.asp"-->
        <div class="container">
            <form class="form-booklist" id="form-booklist" method="post">
                <input type="hidden" name="sortid" id="sortid" value="<%=lngSortId%>" />
                <div class="row" style="margin-bottom:10px;">
                    <div class="col-xs-12">
                        <div class="input-group">
                            <input type="text" class="form-control" name="bookKeyWords" value="<%=strKeyWords%>" id="bookKeyWords" placeholder="输入关键字...">
                            <span class="input-group-btn">
                                <button class="btn btn-primary" type="button" id="booklist-search-btn"><span class="glyphicon glyphicon-search" aria-hidden="true"></span> 搜索</button>
                            </span>
                        </div>
                    </div>
                </div>
            </form>

            <div id="gvBooks" class="row">

            </div>
        </div>
        <!-- /container -->
        <!--#include file="footer.asp"-->
        <script id="tplItem" type="text/x-jsrender">
            <div class="col-xs-6">
                <div class="thumbnail">
                    <img src="<%=gstrInstallDir%>uppic/big/{{:picurl}}" alt="{{:title}}" class="img-responsive img-rounded shopbook-img" />
                    <h4 class="text-center">{{:title}}</h4>
                    <p class="text-right"><button class="btn btn-success btn-xs fav-btn" infoid="{{:infoId}}" role="button"><span class="glyphicon glyphicon-star" aria-hidden="true"></span> 收藏</button></p>
                </div>
            </div>
        </script>
        <div class="col-xs-6">
            <div class="thumbnail">
                <img src="<%=gstrInstallDir%>uppic/big/aa.jpg alt="ddd" class="img-responsive img-rounded shopbook-img" />
                <h4 class="text-center">dddddd</h4>
                <p class="text-right"><button class="btn btn-success btn-xs fav-btn" infoid="111" role="button"><span class="glyphicon glyphicon-star" aria-hidden="true"></span> 收藏</button></p>
            </div>
        </div>
        <script src="//cdn.bootcss.com/jquery/3.2.1/jquery.min.js"></script>
        <script src="//cdn.bootcss.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
        <script src="//cdn.bootcss.com/jsrender/1.0.0-rc.70/jsrender.min.js"></script>
        <script>

            $(function () {
            
                //收藏
                $(".fav-btn").on("click", function(){
                    //AJAX提交下面地址
                    //servece/bookfav-insert.asp
                    alert($(this).attr("infoid"));
                });

            });

        </script>
    </body>
</html>