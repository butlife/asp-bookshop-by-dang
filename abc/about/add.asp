<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>Js/common.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
function Body_Load(){
}

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
<body onLoad="Body_Load();">
<form name="form1" method="post" action="Save.asp" onSubmit="return btnSubmit_Click();">
<input name="savetype" type="hidden" value="add">
  <div id="headPanel">站点信息添加</div>
  <div id="bodyContent">
    <div>
      <label id="ltitle">标题</label>
      <input id="title" name="title" type="text" size="40" maxlength="48" />
    </div>
    <div>
      <label id="lblContent">内容</label>
	  <textarea name="s_News" id="s_News" style="width:700px;height:300px;"></textarea>
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
