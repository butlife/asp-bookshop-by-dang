<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
	Dim lngInfoId, strTitle,  strAdminName, lngAdminId, dtUpdateTime, strContent, sRemark, strpicurl
	Dim strSql, rsInfo
	
	lngInfoId = ConvertLong(Request("id") & "")
	Set rsInfo = Server.CreateObject("ADODB.RecordSet")
	strSql = " Select * From about_v where id = " & lngInfoId 
	If rsInfo.State = 1 Then rsInfo.Close
	rsInfo.Open strSql, conn, 1, 1
	If (rsInfo.Bof Or rsInfo.Eof) Then
		Response.write "<script>alert('��վ����Ϣ�����ڻ򼺱�ɾ��.'); history.back();</script>"
		Response.End()
	Else
		strTitle = Trim(rsInfo("Title") & "")
		sRemark = Trim(rsInfo("Remark") & "")
		strcontent = trim(rsInfo("content") & "")
		strpicurl = Trim(rsInfo("picurl") & "")
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
	
	if (formMain.title.value == ""){
		alert(ltitle.innerText + "����Ϊ�գ�");
		formMain.title.focus();
		return false;
	}
	return true;
}
</script>
</head>
<body onLoad="Body_Load();">
<form name="form1" method="post" action="Save.asp" onSubmit="return btnSubmit_Click();">
<input name="savetype" type="hidden" value="modify">
<input name="id" type="hidden" value="<%=lngInfoId%>">
  <div id="headPanel">վ����Ϣ�޸�</div>
  <div id="bodyContent">
    <div>
      <label id="ltitle">����</label>
      <input id="title" name="title" type="text" size="50" maxlength="48"  value="<%=strTitle%>" />
    </div>
    <div style="display:none;">
      <label id="lpicurl">ͼƬ</label>
      <input id="picurl" name="picurl" type="text" size="40" maxlength="48" value="<%=strpicurl%>" />
		<button onClick="imageupload('picurl');">�ϴ�ͼƬ</button>
    </div>
    <div>
      <label id="lblContent">����</label>
	  <textarea name="s_News" id="s_News" style="width:700px;height:300px;"><%= strcontent%></textarea>
    </div>
	<div id="SubPanel">
      <input id="btnSubmit" name="btnSubmit" type="submit" class="Button" value="�� ��" />
      <input id="btnReset" name="btnReset" type="reset" class="Button" value="�� ��" />
    </div>
  </div>
</form>
</body>
</html>
<%
	Call CloseConn()
%>
