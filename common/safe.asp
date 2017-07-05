<%
	call safeChk()
	
	function safeChk()
		If trim(Session("userid")&"") = "0" or Session("useracc") = "" Then
			Call CloseConn
			Response.Write "<script language=""javascript"" type=""text/javascript"">alert(""网络超时或您还没有登陆！"");parent.location.href=""" & gstrInstallDir & "index.asp"";/*NONELOGIN*/</script>"
			Response.End
		End If
	end function
	
	function getUserInfo(op)
		dim rsUser, strsql, strRe
		op = trim(op&"")
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		strsql = "select " & op & " from user_t where UserId = " & ConvertLong(Session("userid")&"")
		if rsUser.state = 1 then rsUser.close
		rsuser.open strsql,conn,1,1
		if not(rsUser.bof or rsUser.eof) then
			strRe = trim(rsuser(op) & "")
		else
			strRe = "0"
		end if
		if rsUser.state = 1 then rsUser.close
		set rsUser = nothing
		getUserInfo = strRe		
	end function
	
	function getUserShopCount()
		dim rsUser, strsql, strRe
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		strsql = "select count(*) as iii from shop_t where shopstate <= 1 and UserId = " & ConvertLong(Session("userid")&"")
		if rsUser.state = 1 then rsUser.close
		rsuser.open strsql,conn,1,1
		if not(rsUser.bof or rsUser.eof) then
			strRe = trim(rsuser("iii") & "")
		end if
		if rsUser.state = 1 then rsUser.close
		set rsUser = nothing
		getUserShopCount = strRe		
	end function
	
	function getusershopisCount(arrId, thisCount)
		thisCount = ConvertLong(thisCount&"")
		dim rsUser, strsql, strRe, ii
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		strsql = "select icount from info_t where ispassed = 1 and infoid in (" & arrId & ")"
		if rsUser.state = 1 then rsUser.close
		rsuser.open strsql,conn,1,1
		ii=0
		do while not(rsUser.bof or rsUser.eof)
			if ConvertLong(rsUser("icount") & "")>0 then
				ii=ii+1
			end if
			rsUser.movenext
		loop
'		response.write ii & "-"
'		response.write thisCount
		if ii<thisCount then
			strRe = 0
		else
			strRe = 1
		end if
		if rsUser.state = 1 then rsUser.close
		set rsUser = nothing
		getusershopisCount = strRe		
	end function
%>