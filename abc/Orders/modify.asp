<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/message.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!--#include file="infosort.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
	Dim lngInfoId, lngsortid, strcontent, ispassed, istop, shit, strauthor, strtitle, iorder, strpicurl, strremark, dtmakedate
	Dim strSql, rsInfo
	
	lngInfoId = ConvertLong(Request("id") & "")
	Set rsInfo = Server.CreateObject("ADODB.RecordSet")
	strSql = " Select * From info_v where infoid = " & lngInfoId 
	If rsInfo.State = 1 Then rsInfo.Close
	rsInfo.Open strSql, conn, 1, 1
	If (rsInfo.Bof Or rsInfo.Eof) Then
		Response.write "<script>alert('����Ϣ�����ڻ򼺱�ɾ��.'); history.back();</script>"
		Response.End()
	Else
		strtitle = trim(rsInfo("title") & "")
		lngsortid = ConvertLong(rsInfo("sortid") & "")
		strcontent = trim(rsInfo("content") & "")
		strremark = trim(rsInfo("remark") & "")
		istop = ConvertLong(rsInfo("istop") & "")
		shit = ConvertLong(rsInfo("hit") & "")
		ispassed = ConvertLong(rsInfo("ispassed") & "")
		strauthor = trim(rsInfo("author") & "")
		iorder = ConvertLong(rsInfo("iorder") & "")
		strpicurl = trim(rsInfo("picurl") & "")
		dtmakedate = trim(rsInfo("makedate") & "")
	End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../../webkind/index.asp"-->
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>Js/common.js" type="text/javascript"></script>

<script language="javascript" type="text/javascript">
function btnSubmit_Click(){
	var formMain = document.forms[0];

	if (formMain.sortid.value == ""){
		alert(lsort.innerText + "����Ϊ�գ�");
		formMain.sortid.focus();
		return false;
	}
	
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
<input name="savetype" type="hidden" value="modify">
<input name="infoid" type="hidden" value="<%=lngInfoId%>">
  <div id="headPanel">�޸���Ϣ</div>
  <div id="bodyContent">
    <div>
      <label id="lsort">���</label>
      <%= SortSelect(lngsortid)%>
    </div>
    <div>
      <label id="ltitle">����</label>
      <input id="title" name="title" type="text" size="50" maxlength="49" value="<%= strtitle%>" />
    </div>
    <div>
		<label id="limageurl">ͼƬ</label>
		<input name="picurl" type="text" id="picurl" size="20" value="<%= strpicurl%>" />
		<iframe height="20" style="margin:auto;" scrolling="no" frameborder="0" width="580" src="<%= gstrInstallDir%>upload/select.asp?frm=form1&frminput=picurl&m=<%=now()%>"></iframe>
    </div>
    <div>
      <label id="lblLock">ѡ��</label>
      <span>���<input id="ispassed" name="ispassed" type="checkbox" value="1" <%if ispassed = 1 then response.write "checked=""true"""%> /></span>
      <span>�ö�<input id="istop" name="istop" type="checkbox" value="1" <%if istop = 1 then response.write "checked=""true"""%> /></span>
      <span>�����<input id="iorder" name="iorder" type="text" size="5" maxlength="8" value="<%= iorder%>" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div>
      <label id="makedate">ʱ��</label>
      <input id="makedate" name="makedate" type="text" size="20" value="<%=dtmakedate%>" maxlength="20" />
    </div>
    <div>
      <label id="lblContent">����</label>
	  <textarea name="s_News" id="s_News" style="width:700px;height:300px;visibility:hidden;"><%= strcontent%></textarea>
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
