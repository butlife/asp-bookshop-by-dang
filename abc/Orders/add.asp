<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!--#include file="infosort.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../../webkind/index.asp"-->
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>Js/common.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
function btnSubmit_Click(){
	var formMain = document.forms[0];

	if (formMain.sortid.value == ""){
		alert(lsort.innerText + "不能为空！");
		formMain.sortid.focus();
		return false;
	}
	
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
  <div id="headPanel">添加信息</div>
  <div id="bodyContent">
    <div>
      <label id="lsort">类别</label>
      <%= SortSelect(0)%>
    </div>
    <div>
      <label id="ltitle">标题</label>
      <input id="title" name="title" type="text" size="50" maxlength="49" />
    </div>
    <div>
		<label id="limageurl">图片</label>
		<input name="picurl" type="text" size="20" id="picurl" />
		<iframe height="20" style="margin:auto;" scrolling="no" frameborder="0" width="580" src="<%= gstrInstallDir%>upload/select.asp?frm=form1&frminput=picurl&m=<%=now()%>"></iframe>
    </div>
    <div>
      <label id="lblLock">选项</label>
      <span>审核<input id="ispassed" name="ispassed" type="checkbox" value="1" checked="true" /></span>
      <span>置顶<input id="istop" name="istop" type="checkbox" value="1" /></span>
      <span>点击量<input id="iorder" name="iorder" type="text" size="5" maxlength="8" value="0" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div>
      <label id="makedate">时间</label>
      <input id="makedate" name="makedate" type="text" size="20" maxlength="20" value="<%=now()%>" />
    </div>
    <div>
      <label id="lblContent">内容</label>
	  <textarea name="s_News" id="s_News" style="width:700px;height:300px;visibility:hidden;"></textarea>
    </div>
    <div id="SubPanel">
      <input id="btnSubmit" name="btnSubmit" type="submit" class="Button" value="保 存" />
      <input id="btnReset" name="btnReset" type="reset" class="Button" value="重 设" />
    </div>
  </div>
</form>
</body>
</html>
<%
	Call CloseConn()
%>
