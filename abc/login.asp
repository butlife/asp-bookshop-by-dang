<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!-- #include file="../common/conn.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%= gstrSiteName%> -- 后台登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="<%= gstrInstallDir%>Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>Js/common.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
function btnSubmit_Click(){
	var frm = document.forms[0];
	if (frm.adminname.value == ""){ 
		alert("请输入您的帐号!"); 
		return false; 
	}
	
	if (frm.adminpwd.value == ""){ 
		alert("请输入您的密码!"); 
		return false; 
	}
	return true;
}
</script>
<style type="text/css">
body { background:#417bc9;}
.white {color: #FFFFFF}
</style>
</head>
<body>
<form id="form1" name="form1" method="post" action="CheckLogon.asp" onSubmit="return btnSubmit_Click();">
<div style="text-align:center; padding:120px; margin:20px; margin:0 auto; border:1px solid #417bc9; background-color:#417bc9; height:300px;">
	<table width="432" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <th height="26" colspan="2" background="images/Login_Top.gif" scope="col"><span class="white"><%= gstrSiteName%> - 后台管理</span></th>
      </tr>
      <tr>
        <td bgcolor="#C4D0E1"><img src="images/Login_TT.jpg" width="132" height="130" hspace="10" vspace="10"></td>
        <td background="images/Login_BG.gif"><table border="0" cellspacing="3" cellpadding="3">
          <tr>
            <td>用户名：</td>
            <td><input name="adminname" type="text" id="adminname" class="input_text" size="16" maxlength="49"></td>
          </tr>
          <tr>
            <td>密 码：</td>
            <td><input name="adminpwd" type="password" id="adminpwd" class="input_text" size="16" maxlength="49"></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" class="button_white" value="登 录">
                <input type="reset" name="reset" class="button_white" value="清 空"></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="36" colspan="2" background="images/Login_Down.gif" align="right" style="padding-right:20px;"><%= gstrSiteName%></td>
      </tr>
    </table>
  </div>
</form>
</body>
</html>
