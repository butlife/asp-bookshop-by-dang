<%
	'============================================
	'��������Pagination
	'��  �ܣ���ʾ��ҳЧ��
	'��  ����strQuery:      ��ѯ����
	'        lngPageCount:  ��ҳ��
	'        lngCurrPage:   ��ǰҳ��
	'        lngPageSize:   ÿҳ������
	'����ֵ����ҳ��Html����
	'============================================
	Function Pagination(strQuery, lngPageCount, lngCurrPage, lngPageSize)
		Dim strHtml, i
		Dim strActionUrl 
		strActionUrl = Trim(Request.ServerVariables("PATH_INFO"))
		
		strHtml = ""
		strHtml = strHtml & "<script language=""javascript"">" & vbCrLf
		strHtml = strHtml & "function GotoPage(iPage){" & vbCrLf
		strHtml = strHtml & "	var formMain = document.forms[0];" & vbCrLf
		strHtml = strHtml & "	formMain.Page.value = iPage;" & vbCrLf
		strHtml = strHtml & "	formMain.action = """ & strActionUrl & """" & vbCrLf
		strHtml = strHtml & "	formMain.submit();" & vbCrLf
		strHtml = strHtml & "}" & vbCrLf
		strHtml = strHtml & "</script>" & vbCrLf
		strHtml = strHtml & "" & vbCrLf
		strHtml = strHtml & "<table cellspacing=""0"" cellfillding=""0"" align=""center"">" & vbCrLf
		strHtml = strHtml & "<input type=""hidden"" name=""Query"" value=""" & strQuery & """>" & vbCrLf
		strHtml = strHtml & "<input type=""hidden"" name=""Page"" value="""">" & vbCrLf
		strHtml = strHtml & "<tr class=""ListPager"">" & vbCrLf
		strHtml = strHtml & "<td align=""left"">" & vbCrLf
		strHtml = strHtml & "&nbsp;&nbsp;ҳ��:" & lngCurrPage & "/" & lngPageCount & "&nbsp;"
		strHtml = strHtml & "��" & lngRecordCount & "&nbsp;ÿ" & lngPageSize & "/ҳ"& vbCrLf
		strHtml = strHtml & "</td>" & vbCrLf
		strHtml = strHtml & "<td align=""right"">" & vbCrLf
		If lngCurrPage = 1 Then
			strHtml = strHtml & "&nbsp;[��ҳ]&nbsp;[��ҳ]"
		Else
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(1)"">[��ҳ]</a>"
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(" & lngCurrPage - 1 & ")"">[��ҳ]</a>"
		End If
		If lngCurrPage = lngPageCount Then
			strHtml = strHtml & "&nbsp;[��ҳ]&nbsp;[βҳ]&nbsp;"
		Else
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(" & lngCurrPage + 1 & ")"">[��ҳ]</a>"
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(" & lngPageCount & ")"">[βҳ]</a>"
			strHtml = strHtml & "&nbsp;" & vbCrLf
		End If
		strHtml = strHtml & "</td>" & vbCrLf
		strHtml = strHtml & "<td align=""left"">" & vbCrLf
		strHtml = strHtml & "<span>ת��</span>" & vbCrLf
		strHtml = strHtml & "<select name=""selPageIndex"" id=""selPageIndex"" onchange=""GotoPage(this.value);"">" & vbCrLf
		For i = 1 To lngPageCount
			strHtml = strHtml & "<option value="""& i & """"
			If i = lngCurrPage Then
				strHtml = strHtml & " selected=""selected"""
			End If
			strHtml = strHtml & ">��"& i & "ҳ</option>" & vbCrLf
		Next
		strHtml = strHtml & "</select> " & vbCrLf
		strHtml = strHtml & "</td>" & vbCrLf
		strHtml = strHtml & "</tr>" & vbCrLf
		strHtml = strHtml & "</table>"
		
		Pagination = strHtml
		
	End Function
	
%>
