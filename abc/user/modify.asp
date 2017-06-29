<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%Response.Charset = "GB2312"%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
	Dim lngUserId, struserName, struserAcc, struserAdd, struserSex, struserpwd, strusertel, strexpdate_s, strexpdate_e, strRemark, lnguseCounts, lngmaxuseCounts, strhobby, lngispassed
	Dim strSql, rsUser
	
	lngUserId = ConvertLong(Request("Id") & "")
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	strSql = " Select * From user_t where userid = " & lnguserId 
	If rsUser.State = 1 Then rsUser.Close
	rsUser.Open strSql, conn, 1, 1
	If (rsUser.Bof Or rsUser.Eof) Then
		Response.write "<script>alert('该会员不存在或己被删除.'); history.back();</script>"
		Response.End()
	Else
		strusername = Trim(rsUser("username") & "")
		struserAcc = Trim(rsUser("userAcc") & "")
		struserAdd = Trim(rsUser("userAdd") & "")
		struserSex = Trim(rsUser("userSex") & "")
		struserpwd = Trim(rsUser("userpwd") & "")
		strusertel = Trim(rsUser("usertel") & "")
		strexpdate_s = Trim(rsUser("expdate_s") & "")
		strexpdate_e = Trim(rsUser("expdate_e") & "")
		lnguseCounts = ConvertLong(rsUser("useCounts") & "")
		lngmaxuseCounts = ConvertLong(rsUser("maxuseCounts") & "")
		strhobby = Trim(rsUser("hobby") & "")
		lngispassed = ConvertLong(rsUser("ispassed") & "")
		strRemark = Trim(rsUser("Remark") & "")
	End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>Js/common.js" type="text/javascript"></script>

<script language="javascript" type="text/javascript">
function Body_Load(){

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
  <div id="headPanel">修改会员</div>
<input name="savetype" type="hidden" value="modify">
<input name="userid" type="hidden" value="<%=lnguserid%>">
  <div id="bodyContent">
    <div>
      <label id="lusername">会员姓名</label>
      <input id="username" name="username" type="text" size="20" maxlength="20" value="<%=strUserName%>" />
    </div>
    <div>
      <label id="lusersex">会员性别</label>
      <select name="usersex" id="usersex">
      	<option value=""></option>
      	<option value="男" <%if strusersex = "男" then response.write "selected=""selected"""%>>先生</option>
      	<option value="女" <%if strusersex = "女" then response.write "selected=""selected"""%>>女士</option>
      </select>
    </div>
    <div>
      <label id="lblUserAcc">会员帐号</label>
      <input id="useracc" name="useracc" type="text" size="20" maxlength="20" value="<%=struseracc%>" readonly="readonly" style="background:#ccc;" />
    </div>
    <div>
      <label id="luserpwd">会员密码</label>
      <input id="userpwd" name="userpwd" type="text" size="20" maxlength="20" value="<%=struserpwd%>" />
    </div>
    <div>
      <label id="lusertel">会员手机</label>
      <input id="usertel" name="usertel" type="text" size="20" maxlength="20" value="<%=strusertel%>" />
    </div>
    <div>
      <label id="luserAdd">送书地址</label>
      <textarea id="userAdd" name="userAdd" cols="68" rows="3" wrap="virtual"><%=struserAdd%></textarea>
    </div>
    <div>
      <label id="lsortname">借书限制</label>
      <span>单次借书数量<input id="useCounts" name="useCounts" type="text" size="5" maxlength="8" value="<%=lnguseCounts%>" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
      <span>最大借书次数<input id="maxuseCounts" name="maxuseCounts" type="text" size="5" maxlength="8" value="<%=lngmaxuseCounts%>" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div>
      <label id="lblLock">是否启用</label>
      <span>审核<input id="ispassed" name="ispassed" type="checkbox" value="1" <%if lngispassed = 1 then response.write "checked=""checked"""%>" /></span>
    </div>
    <div>
      <label id="lblqx">会员期限</label>
      <span>从 <input id="expdate_s" name="expdate_s" type="text" size="12" value="<%=Format_Time(strexpdate_s,2)%>" /> 到 <input id="expdate_e" name="expdate_e" type="text" value="<%=Format_Time(strexpdate_e,2)%>" /> 止</span>
    </div>
    <div>
      <label id="lhobby">兴趣爱好</label>
      <textarea id="hobby" name="hobby" cols="68" rows="3" wrap="virtual"><%=strhobby%></textarea>
    </div>
    <div>
      <label id="lblRemark">其它备注</label>
      <textarea id="Remark" name="Remark" cols="68" rows="3" wrap="virtual"><%=strRemark%></textarea>
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
