<%
	Function infosort_NL(lngSortID)
		Dim rsSort, strSql
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_NL Order By iorder, SortId desc "
		rsSort.Open strSql, conn, 1, 1
		Response.Write "<select id=""infosort_NL_ID"" name=""infosort_NL_ID"" class=""listPicker"">"
		Dim strSortName
		
		Response.Write "<option value=""0""" 
		If lngSortID = "0" Then
			Response.Write "selected=""selected"""
		End If
		Response.Write ">&nbsp;</option>"
		Do While Not (rsSort.Eof)
			strSortName = Trim(rsSort("SortName") & "")
			Response.Write "<option value=""" & rsSort("SortId") & """" 
			If rsSort("SortId") = lngSortID Then
				Response.Write "selected=""selected"""
			End If
			
			Response.Write ">" & strSortName & "</option>"
			rsSort.MoveNext
		Loop
		Response.Write "</select>"
	End Function
	
		Function infosort_ZT(lngSortID)
		Dim rsSort, strSql
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_ZT Order By iorder, SortId desc "
		rsSort.Open strSql, conn, 1, 1
		Response.Write "<select id=""infosort_ZT_ID"" name=""infosort_ZT_ID"" class=""listPicker"">"
		Dim strSortName
		
		Response.Write "<option value=""0""" 
		If lngSortID = "0" Then
			Response.Write "selected=""selected"""
		End If
		Response.Write ">&nbsp;</option>"
		Do While Not (rsSort.Eof)
			strSortName = Trim(rsSort("SortName") & "")
			Response.Write "<option value=""" & rsSort("SortId") & """" 
			If rsSort("SortId") = lngSortID Then
				Response.Write "selected=""selected"""
			End If
			
			Response.Write ">" & strSortName & "</option>"
			rsSort.MoveNext
		Loop
		Response.Write "</select>"
	End Function

	Function infosort_XL(lngSortID)
		Dim rsSort, strSql
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_XL Order By iorder, SortId desc "
		rsSort.Open strSql, conn, 1, 1
		Response.Write "<select id=""infosort_XL_ID"" name=""infosort_XL_ID"" class=""listPicker"">"
		Dim strSortName
		
		Response.Write "<option value=""0""" 
		If lngSortID = "0" Then
			Response.Write "selected=""selected"""
		End If
		Response.Write ">&nbsp;</option>"
		Do While Not (rsSort.Eof)
			strSortName = Trim(rsSort("SortName") & "")
			Response.Write "<option value=""" & rsSort("SortId") & """" 
			If rsSort("SortId") = lngSortID Then
				Response.Write "selected=""selected"""
			End If
			
			Response.Write ">" & strSortName & "</option>"
			rsSort.MoveNext
		Loop
		Response.Write "</select>"
	End Function

	Function infosort_FM(lngSortID)
		Dim rsSort, strSql
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_FM Order By iorder, SortId desc "
		rsSort.Open strSql, conn, 1, 1
		Response.Write "<select id=""infosort_FM_ID"" name=""infosort_FM_ID"" class=""listPicker"">"
		Dim strSortName
		
		Response.Write "<option value=""0""" 
		If lngSortID = "0" Then
			Response.Write "selected=""selected"""
		End If
		Response.Write ">&nbsp;</option>"
		Do While Not (rsSort.Eof)
			strSortName = Trim(rsSort("SortName") & "")
			Response.Write "<option value=""" & rsSort("SortId") & """" 
			If rsSort("SortId") = lngSortID Then
				Response.Write "selected=""selected"""
			End If
			
			Response.Write ">" & strSortName & "</option>"
			rsSort.MoveNext
		Loop
		Response.Write "</select>"
	End Function

	Function infosort_CBS(lngSortID)
		Dim rsSort, strSql
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_CBS Order By iorder, SortId desc "
		rsSort.Open strSql, conn, 1, 1
		Response.Write "<select id=""infosort_CBS_ID"" name=""infosort_CBS_ID"" class=""listPicker"">"
		Dim strSortName
		
		Response.Write "<option value=""0""" 
		If lngSortID = "0" Then
			Response.Write "selected=""selected"""
		End If
		Response.Write ">&nbsp;</option>"
		Do While Not (rsSort.Eof)
			strSortName = Trim(rsSort("SortName") & "")
			Response.Write "<option value=""" & rsSort("SortId") & """" 
			If rsSort("SortId") = lngSortID Then
				Response.Write "selected=""selected"""
			End If
			
			Response.Write ">" & strSortName & "</option>"
			rsSort.MoveNext
		Loop
		Response.Write "</select>"
	End Function

%>