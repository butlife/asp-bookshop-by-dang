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
	
	if (formMain.adminname.value == ""){
		alert(lblAdminName.innerText + "����Ϊ�գ�");
		formMain.adminname.focus();
		return false;
	}

	if (formMain.AdminPwd.value == ""){
		alert(lblAdminPwd.innerText + "����Ϊ�գ�");
		formMain.tbAdminPwd.focus();
		return false;
	}

	if (formMain.AdminPwd.value != formMain.AdminPwd2.value){
		alert(lblAdminPwd.innerText + "��һ�£�");
		formMain.AdminPwd.focus();
		return false;
	}

	formMain.submit();
}

function btnCheckName_Click(){
	var formMain = document.forms[0];

	var strAdminName = formMain.adminname.value;
	
	if(strAdminName == ""){
		alert(lblAdminName.innerText + "����Ϊ�գ�");
		formMain.AdminName.focus();
		return false;
	}
	
	window.open('CheckName.asp?AdminName=' + strAdminName, '_blank', 'status=no,top=100,left=100,width=500,height=250,scrollbars=no');
}
</script>
</head>
<body onLoad="Body_Load();">
<form name="form1" method="post" action="Save.asp">
<input name="savetype" type="hidden" value="add">
  <div id="headPanel">��ӹ���Ա</div>
  <div id="bodyContent">
    <div>
      <label id="lblAdminName">�ʺ�</label>
      <input id="adminname" name="adminname" type="text" size="20" maxlength="20" />
      <input id="btnCheckName" name="btnCheckName" type="button" class="Buttion"  onclick="btnCheckName_Click();" value=" ����ʺ� " />
    </div>
    <div>
      <label id="lblAdminPwd">����</label>
      <input id="AdminPwd" name="AdminPwd" type="password" size="20" maxlength="17" />
    </div>
    <div>
      <label id="lblAdminPwd2">����</label>
      <input id="AdminPwd2" name="AdminPwd2" type="password" size="20" maxlength="17" />
    </div>
    <div>
      <label id="lblLock">��ͨ</label>
      <input id="ckbLock" name="ckbLock" type="checkbox" value="1" />
    </div>
    <div>
      <label id="lblRemark">��ע</label>
      <textarea id="Remark" name="Remark" cols="40" rows="3" wrap="virtual"></textarea>
    </div>
    <div id="SubPanel">
      <input id="btnSubmit" name="btnSubmit" type="button" class="Button" value="�� ��" onClick="btnSubmit_Click();" />
      <input id="btnReset" name="btnReset" type="reset" class="Button" value="�� ��" />
    </div>
  </div>
</form>
</body>
</html>
<%
	Call CloseConn()
%>
