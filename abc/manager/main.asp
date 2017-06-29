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
</script>
</head>
<body onLoad="Body_Load();">
<form action="" method="post" name="form1">
  <div id="headPanel">����Ա�б�</div>
  <div id="buttonPanel">
    <a class="aButton" href="add.asp">����Ա���</a>
  </div>
  <div id="contentPanel">
    <div id="tableContainer" class="tableContainer">
      <table id="fixedList" border="1" cellfillding="0" cellspacing="0" width="98%" >
        <thead class="fixedHeader">
          <tr>
            <th nowrap="nowrap">���</th>
            <th nowrap="nowrap">����</th>
            <th nowrap="nowrap">״̬</th>
            <th nowrap="nowrap">���ʱ��</th>
            <th nowrap="nowrap">�ϴε�½ʱ��</th>
            <th nowrap="nowrap">�ϴε�½IP</th>
            <th nowrap="nowrap">��½����</th>
            <th nowrap="nowrap">����</th>
          </tr>
        </thead>
        <tbody class="scrollContent">
          <%
		'��ȡ���ݣ���ʾ�б�
		Dim lngAdminId, strAdminName, bispassed, dtAddTime, strLastLoginIp, dtlogindate, lngLoginCount, strState

		Dim rsAdmin, i, strSql, strQuery
		Dim lngPageSize, lngPageCount, lngCurrPage, lngRecordCount, bPagination
		
		Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From manager_t Order By id "
		rsAdmin.Open strSql, conn, 1, 1
		
		If Not (rsAdmin.Bof Or rsAdmin.Eof) Then
			bPagination = True
			'��ҳ
			lngPageSize = glngPageSize
			lngRecordCount = rsAdmin.RecordCount
			rsAdmin.PageSize = lngPageSize
			lngPageCount = rsAdmin.PageCount
			If ConvertLong(Request("Page") & "") <> 0 Then
				lngCurrPage = CLng(Request("Page") & "")
			Else
				lngCurrPage = 1
			End If
			If lngCurrPage <= 1 Then lngCurrPage = 1
			If lngCurrPage >= lngPageCount Then lngCurrPage = lngPageCount
			rsAdmin.AbsolutePage = lngCurrPage
			
			i = 0
			
			'��ʾ�б�
			Do While Not (rsAdmin.Bof Or rsAdmin.Eof) 

				lngAdminId = ConvertLong(rsAdmin("ID") & "")
				strAdminName = ReplaceBadChar(rsAdmin("AdminName") & "")
				bispassed = ConvertDouble(rsAdmin("ispassed") & "")
				dtAddTime = Trim(rsAdmin("adddate") & "")
				strLastLoginIp = Trim(rsAdmin("sip") & "")
				dtlogindate = Trim(rsAdmin("logindate") & "")
				lngLoginCount = ConvertDouble(rsAdmin("logocount") & "")
				
				If i Mod 2 = 0 Then	
					Response.Write "<tr class=""ListItem"">" & vbCrLf
				Else
					Response.Write "<tr class=""ListAlternatingItem"">" & vbCrLf
				End If

				
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & ((lngCurrPage - 1) * lngPageSize + i + 1) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strAdminName) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & getState(bispassed) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdDate(dtAddTime) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(dtlogindate) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdString(strLastLoginIp) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & TdNumeric(lngLoginCount) & "</td>" & vbCrLf
				Response.Write "<td nowrap=""nowrap"" align=""center"">" & vbCrLf
				Response.Write "<a class='admin_btn' href='modify.asp?Id=" & lngAdminId & "'>�޸�</a>" & vbCrLf
				If (bispassed = 1) Then
					Response.Write "<a class='admin_btn' href='action.asp?type=passed&tbId=" & lngAdminId & "'>����</a>" & vbCrLf
				Else
					Response.Write "<a class='admin_btn_b' href='action.asp?type=passed&tbId=" & lngAdminId & "'>���</a>" & vbCrLf
				End If
				Response.Write "<a class='admin_btn' href='action.asp?type=delete&tbId=" & lngAdminId & "'>ɾ��</a>" & vbCrLf
				Response.Write "</td>" & vbCrLf
				
				Response.Write "</tr>" & vbCrLf
				
				i = i + 1
				rsAdmin.MoveNext
				If i >= lngPageSize Then Exit Do
			Loop
		End If
		
		If rsAdmin.State = 1 Then rsAdmin.Close
		Set rsAdmin = Nothing
		
	function getState(i)
		dim strRe
		i = trim(i & "")
		select case i
			case "0"
			strRe = "<span style=""color:red; font-weight:bold;"">&times;</span>"
			
			case "1"
			strRe = "<span style=""color:green; font-weight:bold;"">&radic;</span>"
			
			case else
			strRe = ""
		end select
		getState = strRe
	end function
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
