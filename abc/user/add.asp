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
		alert(lblUserAcc.innerText + "����Ϊ�գ�");
		formMain.UserAcc.focus();
		return false;
	}
	
	window.open('CheckName.asp?userAcc=' + strUserAcc, '_blank', 'status=no,top=100,left=100,width=450,height=240,scrollbars=no');
}

function btnSubmit_Click(){
	var formMain = document.forms[0];
	
	if (formMain.useracc.value == ""){
		alert(lblUserAcc.innerText + "����Ϊ�գ�");
		formMain.useracc.focus();
		return false;
	}

	if (formMain.username.value == ""){
		alert(lusername.innerText + "����Ϊ�գ�");
		formMain.username.focus();
		return false;
	}

	if (formMain.userpwd.value == ""){
		alert(luserpwd.innerText + "����Ϊ�գ�");
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
  <div id="headPanel">��ӻ�Ա</div>
  <div id="bodyContent">
    <div>
      <label id="lusername">��Ա����</label>
      <input id="username" name="username" type="text" size="20" maxlength="20" />
    </div>
    <div>
      <label id="lusersex">��Ա�Ա�</label>
      <select name="usersex" id="usersex">
      	<option value=""></option>
      	<option value="��">����</option>
      	<option value="Ů">Ůʿ</option>
      </select>
    </div>
    <div>
      <label id="lblUserAcc">��Ա�ʺ�</label>
      <input id="useracc" name="useracc" type="text" size="20" maxlength="20" />
      <input id="btnCheckName" name="btnCheckName" type="button" class="Buttion"  onclick="btnCheckName_Click();" value=" ����ʺ� " />
    </div>
    <div>
      <label id="luserpwd">��Ա����</label>
      <input id="userpwd" name="userpwd" type="text" size="20" maxlength="20" />
    </div>
    <div>
      <label id="lusertel">��Ա�ֻ�</label>
      <input id="usertel" name="usertel" type="text" size="20" maxlength="20" />
    </div>
    <div>
      <label id="luserAdd">�����ַ</label>
      <textarea id="userAdd" name="userAdd" cols="68" rows="3" wrap="virtual"></textarea>
    </div>
    <div>
      <label id="lsortname">��������</label>
      <span>���ν�������<input id="useCounts" name="useCounts" type="text" size="5" maxlength="8" value="0" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
      <span>���������<input id="maxuseCounts" name="maxuseCounts" type="text" size="5" maxlength="8" value="0" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div>
      <label id="lblLock">�Ƿ�����</label>
      <span>���<input id="ispassed" name="ispassed" type="checkbox" value="1" checked="true" /></span>
    </div>
    <div>
      <label id="lblqx">��Ա����</label>
      <span>�� <input id="expdate_s" name="expdate_s" type="text" size="12" value="<%=Format_Time(date(),2)%>" /> �� <input id="expdate_e" name="expdate_e" type="text" value="<%=Format_Time(DateAdd("yyyy",1,date()),2)%>" /> ֹ</span>
    </div>
    <div>
      <label id="lhobby">��Ȥ����</label>
      <textarea id="hobby" name="hobby" cols="68" rows="3" wrap="virtual"></textarea>
    </div>
    <div>
      <label id="lblRemark">������ע</label>
      <textarea id="Remark" name="Remark" cols="68" rows="3" wrap="virtual"></textarea>
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
