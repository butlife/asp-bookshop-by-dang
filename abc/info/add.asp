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
		alert(ltitle.innerText + "����Ϊ�գ�");
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
  <div id="headPanel">���ͼ��</div>
  <div id="bodyContent">
    <div class="div">
      <label id="ltitle">ͼ������</label>
      <input id="title" name="title" type="text" size="50" maxlength="49" />
    </div>
    <div class="div">
      <label id="licount">�����</label>
      <input id="icount" name="icount" type="text" size="10" maxlength="3" value="1" />
      <red class="red">�˿�������趨�󲻿����׸ı䣬���������쳣</red>
    </div>
    <div class="div">
      <label id="lsort">�������</label>
      <%= infosort_NL(0)%>
    </div>
    <div class="div">
      <label id="lsort">�������</label>
      <%= infosort_ZT(0)%>
    </div>
    <div class="div">
      <label id="lsort">ϵ�����</label>
      <%= infosort_XL(0)%>
    </div>
    <div class="div">
      <label id="lsort">��ĸ���</label>
      <%= infosort_FM(0)%>
    </div>
    <div class="div">
      <label id="lsort">���������</label>
      <%= infosort_CBS(0)%>
    </div>
    <div class="div">
		<label id="limageurl">ͼƬ</label>
		<input name="picurl" type="text" size="20" id="picurl" />
		<iframe height="20" style="margin:auto;" scrolling="no" frameborder="0" width="580" src="<%= gstrInstallDir%>upload/select.asp?frm=form1&frminput=picurl&m=<%=now()%>"></iframe>
    </div>
    <div class="div">
      <label id="lblLock">ѡ��</label>
      <span>���<input id="ispassed" name="ispassed" type="checkbox" value="1" checked="true" /></span>
      <span>�ö�<input id="istop" name="istop" type="checkbox" value="1" /></span>
      <span>�ȶ�<input id="iorder" name="iorder" type="text" size="5" maxlength="8" value="0" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div class="div" style="display:none;">
      <label id="makedate">ʱ��</label>
      <input id="makedate" name="makedate" type="text" size="20" maxlength="20" value="<%=now()%>" />
    </div>
    <div class="div">
      <label id="lblContent">����</label>
		<script type="text/plain" id="s_News" name="s_News" style="width:720px;height:300px;"></script>
    </div>
    <div class="div" id="SubPanel">
      <input id="btnSubmit" name="btnSubmit" type="submit" class="Button" value="�� ��" />
      <input id="btnReset" name="btnReset" type="reset" class="Button" value="�� ��" />
    </div>
  </div>
</form>
<script type="text/javascript">
	//ʵ�����༭��
	var um = UM.getEditor('s_News');
</script>
</body>
</html>
<%
	Call CloseConn()
%>
