<%
	'============================================
	'过程名：Pagination
	'功  能：显示分页效果
	'参  数：strQuery:      查询条件
	'        lngPageCount:  总页数
	'        lngCurrPage:   当前页数
	'        lngPageSize:   每页总行数
	'返回值：分页的Html代码
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
		strHtml = strHtml & "&nbsp;&nbsp;页次:" & lngCurrPage & "/" & lngPageCount & "&nbsp;"
		strHtml = strHtml & "共" & lngRecordCount & "&nbsp;每" & lngPageSize & "/页"& vbCrLf
		strHtml = strHtml & "</td>" & vbCrLf
		strHtml = strHtml & "<td align=""right"">" & vbCrLf
		If lngCurrPage = 1 Then
			strHtml = strHtml & "&nbsp;[首页]&nbsp;[上页]"
		Else
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(1)"">[首页]</a>"
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(" & lngCurrPage - 1 & ")"">[上页]</a>"
		End If
		If lngCurrPage = lngPageCount Then
			strHtml = strHtml & "&nbsp;[下页]&nbsp;[尾页]&nbsp;"
		Else
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(" & lngCurrPage + 1 & ")"">[下页]</a>"
			strHtml = strHtml & "&nbsp;<a href=""javascript:GotoPage(" & lngPageCount & ")"">[尾页]</a>"
			strHtml = strHtml & "&nbsp;" & vbCrLf
		End If
		strHtml = strHtml & "</td>" & vbCrLf
		strHtml = strHtml & "<td align=""left"">" & vbCrLf
		strHtml = strHtml & "<span>转到</span>" & vbCrLf
		strHtml = strHtml & "<select name=""selPageIndex"" id=""selPageIndex"" onchange=""GotoPage(this.value);"">" & vbCrLf
		For i = 1 To lngPageCount
			strHtml = strHtml & "<option value="""& i & """"
			If i = lngCurrPage Then
				strHtml = strHtml & " selected=""selected"""
			End If
			strHtml = strHtml & ">第"& i & "页</option>" & vbCrLf
		Next
		strHtml = strHtml & "</select> " & vbCrLf
		strHtml = strHtml & "</td>" & vbCrLf
		strHtml = strHtml & "</tr>" & vbCrLf
		strHtml = strHtml & "</table>"
		
		Pagination = strHtml
		
	End Function
	
%>
