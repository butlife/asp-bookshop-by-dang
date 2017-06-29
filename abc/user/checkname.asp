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
</script>
</head>
<body onLoad="Body_Load();" style="padding:0; margin:0;">
<%
	Dim lngAdminId, struserAcc
	Dim rsAdmin, strSql
	
	struserAcc = Trim(Request("userAcc") & "")
	
	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	strSql = "Select * From user_t Where useracc = '" & struserAcc & "'"
	If rsAdmin.State = 1 Then rsAdmin.Close
	rsAdmin.Open strSql, conn, 1, 1
	
	If rsAdmin.Eof Or rsAdmin.Bof Then
		WriteSuccessMsg "恭喜您,此帐号 <font color=red>[ " & struserAcc & " ]</font> 还未使用,可以继续!", False
	Else
		WriteErrorMsg "帐号己存在!"
	End If

	If rsAdmin.State = 1 Then rsAdmin.Close
	Set rsAdmin = Nothing
%>
</body>
</html>
<%
	Call CloseConn()
%>
