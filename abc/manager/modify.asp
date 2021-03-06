<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>

<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
	Dim lngAdminId, strAdminName, strRemark, strIsPassed
	Dim strSql, rsAdmin
	
	lngAdminId = ConvertLong(Request("id") & "")
	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	strSql = " Select * From manager_t where id = " & lngAdminId 
	If rsAdmin.State = 1 Then rsAdmin.Close
	rsAdmin.Open strSql, conn, 1, 1
	If (rsAdmin.Bof Or rsAdmin.Eof) Then
		Response.write "<script>alert('该管理员不存在或己被删除.'); history.back();</script>"
		Response.End()
	Else
		strAdminName = rsAdmin("adminname")
		strRemark = rsAdmin("Remark")
		strIsPassed = ConvertLong(rsAdmin("ispassed") & "")
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
	
	if (formMain.adminname.value == ""){
		alert(lblAdminName.innerText + "不能为空！");
		formMain.adminname.focus();
		return false;
	}

	if (formMain.AdminPwd.value != formMain.AdminPwd2.value){
		alert(lblAdminPwd.innerText + "不一致！");
		formMain.AdminPwd.focus();
		return false;
	}

	formMain.submit();
}

function btnCheckName_Click(){
	var formMain = document.forms[0];

	var strAdminName = formMain.adminname.value;
	
	if(strAdminName == ""){
		alert(lblAdminName.innerText + "不能为空！");
		formMain.AdminName.focus();
		return false;
	}
	
	window.open('CheckName.asp?AdminName=' + strAdminName, '_blank', 'status=no,top=100,left=100,width=500,height=250,scrollbars=no');
}
</script>
</head>
<body onLoad="Body_Load();">
<form name="form1" method="post" action="Save.asp">
<input name="savetype" type="hidden" value="modify">
<input name="adminId" type="hidden" value="<%=lngAdminId%>">
  <div id="headPanel">修改管理员</div>
  <div id="bodyContent">
    <div>
      <label id="lblAdminName">帐号</label>
      <input id="adminname" name="adminname" type="text" size="20" maxlength="20" value="<%=strAdminName%>" readonly="true" />
    </div>
    <div>
      <label id="lblAdminPwd">密码</label>
      <input id="AdminPwd" name="AdminPwd" type="password" size="20" maxlength="17" />
    </div>
    <div>
      <label id="lblAdminPwd2">密码</label>
      <input id="AdminPwd2" name="AdminPwd2" type="password" size="20" maxlength="17" />
    </div>
    <div>
      <label id="lblLock">开通</label>
      <input name="ckbLock" type="checkbox" id="ckbLock" value="1" <%If strIsPassed = 1 Then Response.Write "checked=""checked"""%> />
    </div>
    <div>
      <label id="lblRemark">备注</label>
      <textarea id="Remark" name="Remark" cols="40" rows="3" wrap="virtual"><%= strRemark%></textarea>
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
