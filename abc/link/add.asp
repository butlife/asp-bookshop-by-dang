<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%Response.Charset = "GB2312"%>
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

	formMain.submit();
}

function imageupload(obj) {
	var url = "<%= gstrInstallDir%>upload/select.asp?" + Math.random();
	var vArguments = "";
	var sFeatures = "dialogHeight: 100px; dialogWidth: 350px; edge: Raised; center: Yes; help: No; resizable: No; status: No;";
	var strIMGURL = window.showModalDialog(url, vArguments, sFeatures);	
	if ((strIMGURL == '') || (strIMGURL == null) || (strIMGURL == 'null') || (strIMGURL == 'NULL')) {
		document.getElementById(obj).value = '';
	}
		document.getElementById(obj).value = strIMGURL;
}
</script>
</head>
<body onLoad="Body_Load();">
<form name="form1" method="post" action="Save.asp">
<input name="savetype" type="hidden" value="add">
  <div id="headPanel">添加链接</div>
  <div id="bodyContent">
    <div>
      <label id="ltitle">名称</label>
      <input id="title" name="title" type="text" size="20" maxlength="20" />
    </div>
    <div>
      <label id="lhttpurl">链接</label>
      <input id="httpurl" name="httpurl" type="text" size="40" maxlength="45" value="http://" />
    </div>
	<!--<div>
		<label id="limageurl">图片</label>
		<input name="picurl" id="picurl" type="text" />
		<button onClick="imageupload('picurl');">上传图片</button>
    </div>-->
    <div>
      <label id="lblRemark">备注</label>
      <textarea id="Remark" name="Remark" cols="40" rows="3" wrap="virtual"></textarea>
    </div>
    <div id="SubPanel">
      <input id="btnSubmit" name="btnSubmit" type="button" class="Button" value="保 存" onClick="btnSubmit_Click();" />
      <input id="btnReset" name="btnReset" type="reset" class="Button" value="重 设" />
    </div>
  </div>
</form>
</body>
</html>
<%
	Call CloseConn()
%>
