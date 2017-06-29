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
	Dim lngSortid, strsortname, strRemark, iorder
	Dim strSql, rsSort
	
	lngSortid = ConvertLong(Request("id") & "")
	Set rsSort = Server.CreateObject("ADODB.RecordSet")
	strSql = " Select * From infosort_CBS where sortid = " & lngSortid 
	If rsSort.State = 1 Then rsSort.Close
	rsSort.Open strSql, conn, 1, 1
	If (rsSort.Bof Or rsSort.Eof) Then
		Response.write "<script>alert('该信息类别不存在或己被删除.'); history.back();</script>"
		Response.End()
	Else
		strsortname = Trim(rsSort("sortname") & "")
		iorder = ConvertLong(rsSort("iorder") & "")
		strRemark = Trim(rsSort("Remark") & "")
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
	
	if (formMain.sortname.value == ""){
		alert(lsortname.innerText + "不能为空！");
		formMain.sortname.focus();
		return false;
	}

	formMain.submit();
}
</script>
</head>
<body onLoad="Body_Load();">
<form name="form1" method="post" action="Save.asp">
<input name="savetype" type="hidden" value="modify">
<input name="sortid" type="hidden" value="<%=lngSortid%>">
  <div id="headPanel">修改出版社类别</div>
  <div id="bodyContent">
    <div>
      <label id="lsortname">出版社名称</label>
      <input id="sortname" name="sortname" type="text" size="20" maxlength="20" value="<%=strsortname%>" />
    </div>
    <div>
      <label id="liorder">排列序号</label>
      <input id="iorder" name="iorder" type="text" size="20" maxlength="20" value="<%=iorder%>" />
    </div>
    <div>
      <label id="lblRemark">其它备注</label>
      <textarea id="Remark" name="Remark" cols="40" rows="3" wrap="virtual"><%=strRemark%></textarea>
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
