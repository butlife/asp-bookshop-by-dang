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
    <title>网上书城-书库-
        <%=gstrKeyWords%>
    </title>


    <script src="https://cdn.bootcss.com/vue/2.3.4/vue.min.js"></script>
    <script src="https://cdn.bootcss.com/vue-resource/1.3.4/vue-resource.min.js"></script>

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
            width: 150px;
            height: 150px;
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
                        <img v-bind:src="'<%=gstrInstallDir%>uppic/big/' + item.picurl" alt="{{item.title}}" @click="showbox(index)" class="img-responsive img-rounded shopbook-img" />
                        <h4 class="text-center">{{item.title}}</h4>
                        <p class="text-right">
                            <button class="btn btn-xs fav-btn btn-success" infoid="{{item.infoId}}" role="button" @click="fav_click(index)">
                            <span class="glyphicon glyphicon-star" aria-hidden="true"></span> {{item.fav == '0' ? '收藏' : '己收藏'}}</button>
                        </p>
                    </div>
                </div>
            </div>

             <div class="modal fade" id="modal_box" role="dialog">
                <div class="modal-dilog" role=ducument>
                    <div class="modal-content">
                        <div class="modal-header">
                            <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                            <h4 class="modal-title">{{modalmsg.title}}</h4>
                        </div>
                        <div class="modal-body">
                            <p>{{modalmsg.content}}</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
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
        <script type="text/javascript">
            Vue.use(VueResource);
            Vue.http.options.emulateJSON = true;

            const getStyle = (element, attr, NumberMode = 'int') => {
                let target;
                // scrollTop 获取方式不同，没有它不属于style，而且只有document.body才能用
                if (attr === 'scrollTop') {
                    target = element.scrollTop;
                } else if (element.currentStyle) {
                    target = element.currentStyle[attr];
                } else {
                    target = document.defaultView.getComputedStyle(element, null)[attr];
                }
                //在获取 opactiy 时需要获取小数 parseFloat
                return NumberMode == 'float' ? parseFloat(target) : parseInt(target);
            }

            var modal = $('#modal_box').modal({show:false});

            var app = new Vue({
                el: '#app',
                created: function() {
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
                    modalmsg:{title:'', content:''},
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
                                    } else {
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
                                if (json.state == 0 || true) {
                                    vm.booklist[index].fav = '1';
                                }
                            });
                        });
                    },
                    load_more: function() {
                        if (!this.loading) {
                            if (this.current_page < this.max_page) {
                                this.loading = true;
                                this.current_page++;
                                this.load_book();
                            }
                        }
                    },
                    showbox:function(index){
                        var book = this.booklist[index];
                        this.modalmsg = {title:book.title, content:book.content};
                        modal.modal('show');
                    },
                },
                directives: {
                    'load-more': {
                        bind: (el, binding) => {
                            let windowHeight = window.screen.height;
                            let height;
                            let setTop;
                            let paddingBottom;
                            let marginBottom;
                            let requestFram;
                            let oldScrollTop;
                            let scrollEl;
                            let heightEl;
                            let scrollType = el.attributes.type && el.attributes.type.value;
                            let scrollReduce = 2;
                            if (scrollType == 2) {
                                scrollEl = el;
                                heightEl = el.children[0];
                            } else {
                                scrollEl = document.body;
                                heightEl = el;
                            }

                            el.addEventListener('touchstart', () => {
                                height = heightEl.clientHeight;
                                if (scrollType == 2) {
                                    height = height
                                }
                                setTop = el.offsetTop;
                                paddingBottom = getStyle(el, 'paddingBottom');
                                marginBottom = getStyle(el, 'marginBottom');
                            }, false)

                            el.addEventListener('touchmove', () => {
                                loadMore();
                            }, false)

                            el.addEventListener('touchend', () => {
                                oldScrollTop = scrollEl.scrollTop;
                                moveEnd();
                            }, false)

                            function moveEnd() {
                                requestFram = requestAnimationFrame(() => {
                                    if (scrollEl.scrollTop != oldScrollTop) {
                                        oldScrollTop = scrollEl.scrollTop;
                                        moveEnd()
                                    } else {
                                        cancelAnimationFrame(requestFram);
                                        height = heightEl.clientHeight;
                                        loadMore();
                                    }
                                })
                            }

                            function loadMore() {
                                if (scrollEl.scrollTop + windowHeight >= height + setTop + paddingBottom + marginBottom - scrollReduce) {
                                    binding.value();
                                }
                            }
                        }
                    }
                }
            });
        </script>
    </body>

</html>