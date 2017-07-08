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
    <title>网上书城-书库-<%=gstrKeyWords%></title>
    <!-- Bootstrap -->
    <link rel="stylesheet" href="https://cdn.bootcss.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">
    <link href="<%=gstrInstallDir%>bootstrap337/css/bootstrap.min.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://cdn.bootcss.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
	<![endif]-->
        <style>
            .shopbook-img {
                /*width:100px;
                height:100px;*/
            }

.modal-mask {
  position: fixed;
  z-index: 9998;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background-color: rgba(0, 0, 0, .5);
  display: table;
  transition: opacity .3s ease;
}

.modal-wrapper {
  display: table-cell;
  vertical-align: middle;
}

.modal-container {
  width: 300px;
  margin: 0px auto;
  padding: 20px 30px;
  background-color: #fff;
  border-radius: 2px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, .33);
  transition: all .3s ease;
  font-family: Helvetica, Arial, sans-serif;
}

.modal-header h3 {
  margin-top: 0;
  color: #42b983;
}

.modal-body {
  margin: 20px 0;
}

.modal-default-button {
  float: right;
}

/*
 * The following styles are auto-applied to elements with
 * transition="modal" when their visibility is toggled
 * by Vue.js.
 *
 * You can easily play with the modal transition by editing
 * these styles.
 */

.modal-enter {
  opacity: 0;
}

.modal-leave-active {
  opacity: 0;
}

.modal-enter .modal-container,
.modal-leave-active .modal-container {
  -webkit-transform: scale(1.1);
  transform: scale(1.1);
}
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
        <div class="container" id="app">
            <vmodal :title="modalmsg.title" :text="modalmsg.text" :show="showmodal" v-on:close="close"></vmodal>
            <div class="row" style="margin-bottom:10px;">
                <div class="col-xs-12">
                    <div class="input-group">
                        <input type="text" class="form-control" name="bookKeyWords" v-model='search_text' id="bookKeyWords" placeholder="输入关键字...">
                        <span class="input-group-btn">
                                <button class="btn btn-primary" type="button" id="booklist-search-btn" @click='load_book(true)'>
                                    <span class="glyphicon glyphicon-search" aria-hidden="true"></span> 搜索
                                </button>
                            </span>
                        </div>
                    </div>
                </div>

            <div id="gvBooks" class="row" v-load-more="load_more">
                <div class="col-xs-6" v-for="(item, index) in booklist">
                    <div class="thumbnail">
                        <img v-bind:src="'<%=gstrInstallDir%>uppic/big/' + item.picurl" @click="showbox(index)" class="img-responsive img-rounded shopbook-img" />
                        <h5 class="text-center">{{item.title}}</h5>
                        <p class="text-right">
                            <button class="btn btn-xs fav-btn btn-success" v-bind:infoid="item.infoId" role="button" @click="fav_click(index)">
                            <span class="glyphicon glyphicon-star" aria-hidden="true"></span> {{item.fav == '0' ? '收藏' : '己收藏'}}</button>
                        </p>
                    </div>
                </div>
            </div>
        </div>

         <template id="template-modal">
            <transition name="modal">
    <div class="modal-mask" v-show="show">
      <div class="modal-wrapper">
        <div class="modal-container">
          <div class="modal-header">
            <slot name="header">
              {{title}}
            </slot>
          </div>

          <div class="modal-body">
            <slot name="body">
              {{text}}
            </slot>
          </div>

          <div class="modal-footer">
            <slot name="footer">
              <button class="modal-default-button" @click="close">
                关闭
              </button>
            </slot>
          </div>
        </div>
      </div>
    </div>
  </transition>
            </template> 

        <!-- /container -->
        <!--#include file="footer.asp"-->

        <script src="//cdn.bootcss.com/jquery/3.2.1/jquery.min.js"></script>
        <script src="//cdn.bootcss.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
        <script>
            $(function() {
                $('#nav-userinfo').click(function() {
                    location.href = 'main.asp';
                });
            });
            
        </script>
            <script src="https://cdn.bootcss.com/vue/2.3.4/vue.js"></script>
    <script src="https://cdn.bootcss.com/vue-resource/1.3.4/vue-resource.js"></script>
    <script type="text/javascript" src="<%=gstrInstallDir%>js/vue-load-more.js"></script>
        <script type="text/javascript">
            Vue.use(VueResource);
            Vue.http.options.emulateJSON = true;

            const getStyle = (element, attr, NumberMode = 'int') => {
    let target;
    // scrollTop 获取方式不同，没有它不属于style，而且只有document.body才能用
    if (attr === 'scrollTop') { 
        target = element.scrollTop;
    }else if(element.currentStyle){
        target = element.currentStyle[attr]; 
    }else{ 
        target = document.defaultView.getComputedStyle(element,null)[attr]; 
    }
    //在获取 opactiy 时需要获取小数 parseFloat
    return  NumberMode == 'float'? parseFloat(target) : parseInt(target);
}

        const modal = Vue.component('vmodal', {
            template:'#template-modal',
            props:['show', 'title', 'text'],
            created:function(){
                console.log('vmodal created');
            },
            data:function(){
                return {
                   
                };
            },
            methods:{
                close:function(){
                    this.$emit('close');
                },
            },
        });

            var app = new Vue({
                el:'#app',
                created:function(){
                    console.log('app created');
                    this.load_book();
                },
                data: {
                    booklist: [],
                    sortid: '<%=lngSortId%>',
                    search_text: '<%=strKeyWords%>',
                    current_page: -1,
                    max_page: 10,
                    loading: false,
                    showmodal:false,
                    modalmsg:{title:'', text:''},
                },
                methods: {
                    load_book: function(search) {
                        let vm = this;
                        vm.$http.post('service/bookshop.asp', {
                            PageNum: vm.current_page,
                            bookKeyWords: vm.search_text,
                            PageNum: vm.current_page
                        }).then(function(response) {
                            response.json().then(function(json) {
                                console.log('bookshop.asp', json);
                                if (json.state == 0) {
                                    if (vm.current_page == -1 || search === true) {
                                        vm.booklist = json.body;
                                    }else{
                                        vm.booklist = vm.booklist.concat(json.body);
                                    }
                                }
                                vm.loading = false;
                            });
                        });
                    },
                    fav_click: function(index) {
                        console.log('fav_click', index, this.booklist[index]);
                        let vm = this;
                        vm.$http.post('service/bookfav-insert.asp', {
                            InfoId: vm.booklist[index].infoId,
                            fav: vm.booklist[index].fav
                        }).then(function(response) {
                            response.json().then(function(json) {
                                console.log('bookfav-insert.asp', json);
                                if(json.state == 0 || true){
                                    vm.booklist[index].fav = '1';
                                }
                                if(json.state == 2 || false){
                                    vm.booklist[index].fav = '0';
                                }
                            });
                        });
                    },
                    load_more:function(){
                        if(!this.loading){
                            if(this.current_page < this.max_page){
                                this.loading = true;
                                this.current_page++;
                                this.load_book();
                            }
                        }
                    },
                    showbox:function(index){
                        var book = this.booklist[index];
                        this.modalmsg = {title:book.title, text:book.content};
                        this.showmodal = true;
                        console.log('showbox', index);
                    },
                    close:function(){
                        this.showmodal = false;
                        console.log('modal click close');
                    },
                },
            });

        </script>
    </body>
</html>