<%
	Function SortSelect(lngSortID)
		Dim rsSort, strSql
		Set rsSort = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From infosort_CBS Order By iorder, SortId desc "
		rsSort.Open strSql, conn, 1, 1
		Response.Write "<select id=""sortid"" name=""sortid"" class=""listPicker"">"
		Dim strSortName
		
		Do While Not (rsSort.Eof)
			strSortName = Trim(rsSort("SortName") & "")
			Response.Write "<option value=""" & rsSort("SortId") & """" 
			If rsSort("SortId") = lngSortID Then
				Response.Write "selected=""selected"""
			End If
			
			Response.Write ">" & strSortName & "</option>"
			rsSort.MoveNext
		Loop
	End Function
%>