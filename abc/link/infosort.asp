<%
	Function SortSelect(lngSortIdId)
		Dim rsAds, strSql
		Set rsAds = Server.CreateObject("ADODB.RecordSet")
		strSql = "Select * From link_t Order By iorder, SortIdId desc "
		rsAds.Open strSql, conn, 1, 1
		Response.Write "<select id=""SortIdId"" name=""SortIdId"" class=""listPicker"">"
		Dim strSortName
		
		Do While Not (rsAds.Eof)
			strSortName = Trim(rsAds("SortName") & "")
			Response.Write "<option value=""" & rsAds("SortIdId") & """" 
			If rsAds("SortIdId") = lngSortIdId Then
				Response.Write "selected=""selected"""
			End If
			
			Response.Write ">" & strSortName & "</option>"
			rsAds.MoveNext
		Loop
	End Function
%>