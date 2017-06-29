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
	Dim lngInfoId, strcontent, ispassed, istop, shit, strauthor, strtitle, iorder, strpicurl, strremark, dtmakedate
	Dim strSql, rsInfo
	dim lnginfosort_CBS_ID, lnginfosort_FM_ID, lnginfosort_NL_ID, lnginfosort_XL_ID, lnginfosort_ZT_ID, lngiCount
	
	lngInfoId = ConvertLong(Request("id") & "")
	Set rsInfo = Server.CreateObject("ADODB.RecordSet")
	strSql = " Select * From info_t where infoid = " & lngInfoId 
	If rsInfo.State = 1 Then rsInfo.Close
	rsInfo.Open strSql, conn, 1, 1
	If (rsInfo.Bof Or rsInfo.Eof) Then
		Response.write "<script>alert('该信息不存在或己被删除.'); history.back();</script>"
		Response.End()
	Else
		strtitle = trim(rsInfo("title") & "")
		lngiCount = ConvertLong(rsInfo("iCount") & "")
		lnginfosort_CBS_ID = ConvertLong(rsInfo("infosort_CBS_ID") & "")
		lnginfosort_FM_ID = ConvertLong(rsInfo("infosort_FM_ID") & "")
		lnginfosort_NL_ID = ConvertLong(rsInfo("infosort_NL_ID") & "")
		lnginfosort_XL_ID = ConvertLong(rsInfo("infosort_XL_ID") & "")
		lnginfosort_ZT_ID = ConvertLong(rsInfo("infosort_ZT_ID") & "")
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
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>Js/common.js" type="text/javascript"></script>

<script language="javascript" type="text/javascript">
function btnSubmit_Click(){
	var formMain = document.forms[0];

	if (formMain.title.value == ""){
		alert(ltitle.innerText + "不能为空！");
		formMain.title.focus();
		return false;
	}
	
	if (formMain.icount.value == "") || ((formMain.icount.value == "0")){
		alert(licount.innerText + "不能为空！");
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
  <div id="headPanel">修改信息</div>
  <div id="bodyContent">
    <div>
      <label id="ltitle">图书名称</label>
      <input id="title" name="title" type="text" size="50" maxlength="49" value="<%= strtitle%>" />
    </div>
    <div>
      <label id="licount">库存量</label>
      <input id="icount" name="icount" type="text" size="10" maxlength="3" value="<%= lngiCount%>" />
      <red class="red">此库存量，设定后不可轻易改变，否则会产生异常</red>
    </div>
    <div>
      <label id="lsort">年龄类别</label>
      <%= infosort_NL(lnginfosort_NL_ID)%>
    </div>
    <div>
      <label id="lsort">主题类别</label>
      <%= infosort_ZT(lnginfosort_ZT_ID)%>
    </div>
    <div>
      <label id="lsort">系列类别</label>
      <%= infosort_XL(lnginfosort_XL_ID)%>
    </div>
    <div>
      <label id="lsort">父母类别</label>
      <%= infosort_FM(lnginfosort_FM_ID)%>
    </div>
    <div>
      <label id="lsort">出版社类别</label>
      <%= infosort_CBS(lnginfosort_CBS_ID)%>
    </div>
    <div>
		<label id="limageurl">图片</label>
		<input name="picurl" type="text" id="picurl" size="20" value="<%= strpicurl%>" />
		<iframe height="20" style="margin:auto;" scrolling="no" frameborder="0" width="580" src="<%= gstrInstallDir%>upload/select.asp?frm=form1&frminput=picurl&m=<%=now()%>"></iframe>
    </div>
    <div>
      <label id="lblLock">选项</label>
      <span>审核<input id="ispassed" name="ispassed" type="checkbox" value="1" <%if ispassed = 1 then response.write "checked=""true"""%> /></span>
      <span>置顶<input id="istop" name="istop" type="checkbox" value="1" <%if istop = 1 then response.write "checked=""true"""%> /></span>
      <span>热度<input id="iorder" name="iorder" type="text" size="5" maxlength="8" value="<%= iorder%>" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" /></span>
    </div>
    <div style="display:none;">
      <label id="makedate">时间</label>
      <input id="makedate" name="makedate" type="text" size="20" value="<%=dtmakedate%>" maxlength="20" />
    </div>
    <div>
      <label id="lblContent">内容</label>
	  <textarea name="s_News" id="s_News" style="width:700px;height:300px;"><%= strcontent%></textarea>
    </div>
    <div id="SubPanel">
      <input id="btnSubmit" name="btnSubmit" type="submit" class="Button" value="保 存" />
      <input id="btnReset" name="btnReset" type="reset" class="Button" value="重 设" />
    </div>
  </div>
</form>
</body>
</html>
<%
	Call CloseConn()
%>
