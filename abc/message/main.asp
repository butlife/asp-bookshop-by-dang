<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Common/Conn.asp"-->
<!--#include file="../../Common/Function.asp"-->
<!--#include file="../../Common/Pagination.asp"-->
<!--#include file="../Safety/Safety.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="<%= gstrInstallDir%>Css/Style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%= gstrInstallDir%>js/common.js" type="text/javascript"></script>

<script language="javascript" type="text/javascript">
function Body_Load(){

}

function btnDelete_Click(id){
	if(!confirm("��ȷ��Ҫɾ��������¼��?")){
		return;
	}
	location.href = "action.asp?Type=delete&tbId=" + id;
}

function btnView_Click(id) {
	var url = "view.asp?Id=" + id;
	openWindow(url, 500, 450)
}

function btnMsg_Click(id) {
	var url = "msg.asp?messid=" + id + "&s=" + Math.random();
	openDialog(url, 450, 300)
}

</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">���ԣ�Ͷ�ߣ������б�</div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">���</th>
            <th nowrap="nowrap">�鿴</th>
            <th nowrap="nowrap">����</th>
            <th nowrap="nowrap">�绰</th>
            <!--<th nowrap="nowrap">�ƶ��绰</th>
            <th nowrap="nowrap">״̬</th>-->
            <th nowrap="nowrap">IP</th>
            <th nowrap="nowrap">ʱ��</th>
            <th nowrap="nowrap">����</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
<%
		'��ȡ���ݣ���ʾ�б�
		Dim rsmess, i, strSql, lngregid, strrecontent
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination, strQuery
		Dim strregname, strsex, strage, strtelephone, strremark, strIP, strMakedate, strtype, strstate
		
		Set rsmess = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From message_t Order By regid desc"
		rsmess.Open strSql, conn, 1, 1
		
		If Not (rsmess.Bof Or rsmess.Eof) Then
			bPagination = True
			'��ҳ
			lngPageSize = glngPageSize
			lngRecordCount = rsmess.RecordCount
			rsmess.PageSize = lngPageSize
			lngPageCount = rsmess.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsmess.AbsolutePage = lngCurrPage
			
			i = 0
			
			'��ʾ�б�
			Do While Not (rsmess.Bof Or rsmess.Eof) 
				lngregid = ConvertLong(rsmess("regid") & "")
				strregname = trim(rsmess("regname") & "")
				strtype = trim(rsmess("stype") & "")
				if strtype = "1" then strtypeName = "���ԣ�Ͷ�ߣ�����"
				if strtype = "2" then strtypeName = "��Ʊ���"
				strtelephone = trim(rsmess("telephone") & "")
				strIP = trim(rsmess("IP") & "")
				strMakedate = trim(rsmess("makedate") & "")
				strrecontent = trim(rsmess("recontent") & "")
				strremark = trim(rsmess("remark") & "")
				'strstate = "<span style='color:#0f0;'>���ظ�</span>"
				'if strrecontent = "" then strstate = "<span style='color:#f00;'>���ظ�</span>"
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If
		
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""left""><span class='admin_view' onclick='btnView_Click(" & lngregid & ");'>�鿴</span></td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strregname) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strtelephone) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strIP) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & strMakedate & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<span class='admin_btn' onclick='btnDelete_Click(" & lngregid & ");'>ɾ��</span>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				Response.Write "</tr>" & vbCrLf
				i = i + 1
				rsmess.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsmess.State = 1 Then rsmess.Close
		Set rsmess = Nothing
%>
        </tbody>
      </table>
    </div>
    <div id="PaginationPanel">
<%
		'��ʾ��ҳЧ��
		If bPagination Then
			strQuery = inHTML(strQuery)
			Response.Write Pagination(strQuery, lngPageCount, lngCurrPage, lngPageSize)
		End If
%>
    </div>
  </div>
  </div>
</form>
</body>
</html>
<%
	Call CloseConn()
%>
