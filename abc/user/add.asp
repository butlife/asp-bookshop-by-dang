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

function btnCheckName_Click(){
	var formMain = document.forms[0];

	var strUserAcc = formMain.useracc.value;
	
	if(strUserAcc == ""){
		alert(lblUserAcc.innerText + "不能为空！");
		formMain.UserAcc.focus();
		return false;
	}
	
	window.open('CheckName.asp?userAcc=' + strUserAcc, '_blank', 'status=no,top=100,left=100,width=450,height=240,scrollbars=no');
}

function btnSubmit_Click(){
	var formMain = document.forms[0];
	
	if (formMain.useracc.value == ""){
		alert(lblUserAcc.innerText + "不能为空！");
		formMain.useracc.focus();
		return false;
	}

	if (formMain.username.value == ""){
		alert(lusername.innerText + "不能为空！");
		formMain.username.focus();
		return false;
	}

	if (formMain.userpwd.value == ""){
		alert(luserpwd.innerText + "不能为空！");
		formMain.userpwd.focus();
		return false;
	}

	formMain.submit();
}
</script>
</head>
<body onLoad="Body_Load();">
<form name="form1" method="post" action="Save.asp">
<input name="savetype" type="hidden" value="add">
  <div id="headPanel">添加会员</div>
  <div id="bodyContent">
    <div>
      <label id="lusername">会员姓名</label>
      <input id="username" name="username" type="text" size="20" maxlength="20" />
    </div>
    <div>
      <label id="lusersex">会员性别</label>
      <select name="usersex" id="usersex">
      	<option value=""></option>
      	<option value="男">先生</option>
      	<option value="女">女士</option>
      </select>
    </div>
    <div>
      <label id="lblUserAcc">会员帐号</label>
      <input id="useracc" name="useracc" type="text" size="20" maxlength="20" />
      <input id="btnCheckName" name="btnCheckName" type="button" class="Buttion"  onclick="btnCheckName_Click();" value=" 检测帐号 " />
    </div>
    <div>
      <label id="luserpwd">会员密码</label>
      <input id="userpwd" name="userpwd" type="text" size="20" maxlength="20" />
    </div>
    <div>
      <label id="lusertel">会员手机</label>
      <input id="usertel" name="usertel" type="text" size="20" maxlength="20" />
    </div>
    <div>
      <label id="luserAdd">送书地址</label>
      <textarea id="userAdd" name="userAdd" cols="68" rows="3" wrap="virtual"></textarea>
    </div>
    <div>
      <label id="lsortname">借书限制</label>
      <span>单次借书数量<input id="useCounts" name="useCounts" type="text" size="5" maxlength="8" value="0" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
      <span>最大借书次数<input id="maxuseCounts" name="maxuseCounts" type="text" size="5" maxlength="8" value="0" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div>
      <label id="lblLock">是否启用</label>
      <span>审核<input id="ispassed" name="ispassed" type="checkbox" value="1" checked="true" /></span>
    </div>
    <div>
      <label id="lblqx">会员期限</label>
      <span>从 <input id="expdate_s" name="expdate_s" type="text" size="12" value="<%=Format_Time(date(),2)%>" /> 到 <input id="expdate_e" name="expdate_e" type="text" value="<%=Format_Time(DateAdd("yyyy",1,date()),2)%>" /> 止</span>
    </div>
    <div>
      <label id="lhobby">兴趣爱好</label>
      <textarea id="hobby" name="hobby" cols="68" rows="3" wrap="virtual"></textarea>
    </div>
    <div>
      <label id="lblRemark">其它备注</label>
      <textarea id="Remark" name="Remark" cols="68" rows="3" wrap="virtual"></textarea>
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
