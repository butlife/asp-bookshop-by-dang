<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!--#include file="infosort.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<link href="<%= gstrInstallDir%>Css/Style2.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>Js/common.js" type="text/javascript"></script>
<!--#include file="../../webUM123/index.asp"-->
<script language="javascript" type="text/javascript">
function btnSubmit_Click(){
	var formMain = document.forms[0];
	
	if (formMain.title.value == ""){
		alert(ltitle.innerText + "不能为空！");
		formMain.title.focus();
		return false;
	}
	return true;
}
</script>
</head>
<body>
<form name="form1" method="post" action="Save.asp" onSubmit="return btnSubmit_Click();">
<input name="savetype" type="hidden" value="add">
  <div id="headPanel">添加图书</div>
  <div id="bodyContent">
    <div class="div">
      <label id="ltitle">图书名称</label>
      <input id="title" name="title" type="text" size="50" maxlength="49" />
    </div>
    <div class="div">
      <label id="licount">库存量</label>
      <input id="icount" name="icount" type="text" size="10" maxlength="3" value="1" />
      <red class="red">此库存量，设定后不可轻易改变，否则会产生异常</red>
    </div>
    <div class="div">
      <label id="lsort">年龄类别</label>
      <%= infosort_NL(0)%>
    </div>
    <div class="div">
      <label id="lsort">主题类别</label>
      <%= infosort_ZT(0)%>
    </div>
    <div class="div">
      <label id="lsort">系列类别</label>
      <%= infosort_XL(0)%>
    </div>
    <div class="div">
      <label id="lsort">父母类别</label>
      <%= infosort_FM(0)%>
    </div>
    <div class="div">
      <label id="lsort">出版社类别</label>
      <%= infosort_CBS(0)%>
    </div>
    <div class="div">
		<label id="limageurl">图片</label>
		<input name="picurl" type="text" size="20" id="picurl" />
		<iframe height="20" style="margin:auto;" scrolling="no" frameborder="0" width="580" src="<%= gstrInstallDir%>upload/select.asp?frm=form1&frminput=picurl&m=<%=now()%>"></iframe>
    </div>
    <div class="div">
      <label id="lblLock">选项</label>
      <span>审核<input id="ispassed" name="ispassed" type="checkbox" value="1" checked="true" /></span>
      <span>置顶<input id="istop" name="istop" type="checkbox" value="1" /></span>
      <span>热度<input id="iorder" name="iorder" type="text" size="5" maxlength="8" value="0" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div class="div" style="display:none;">
      <label id="makedate">时间</label>
      <input id="makedate" name="makedate" type="text" size="20" maxlength="20" value="<%=now()%>" />
    </div>
    <div class="div">
      <label id="lblContent">内容</label>
		<script type="text/plain" id="s_News" name="s_News" style="width:720px;height:300px;"></script>
    </div>
    <div class="div" id="SubPanel">
      <input id="btnSubmit" name="btnSubmit" type="submit" class="Button" value="保 存" />
      <input id="btnReset" name="btnReset" type="reset" class="Button" value="重 设" />
    </div>
  </div>
</form>
<script type="text/javascript">
	//实例化编辑器
	var um = UM.getEditor('s_News');
</script>
</body>
</html>
<%
	Call CloseConn()
%>
