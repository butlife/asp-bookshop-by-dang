<!-- #include file="../Config-utf.asp" -->
<%

	Dim dblStartTimer
	dblStartTimer = Timer()
	
	'常量
	Const DATABASE_ACCESS = 1
	Const DATABASE_MSSQL = 2
	
	'页全局变量区
	Dim gstrSiteName, gstrSiteTitle, gstrLogoUrl, gstrBannerUrl, gstrKeyWords, gstrdescription,gstrServiceTel, glngPageSize_phone
	Dim gstrWebmasterName, gstrWebmasterEmail, gstrCopyright
	Dim gstrSiteUrl, gstrInstallDir, glngSessionTimeout, glngPageSize, gstrNoPicUrl, gstrNoFriendLinkPicUrl, gstrNoFriendLink
	Dim gstrPicFolderUrl, gstrFileFolderUrl, gstrNewsFolderUrl, gstrAdminPanelUrl, gstrUserPanelUrl
	Dim gstrFSOName, gstrSessionPrefix
	
	Dim gstrAllowExt, gstrUpLoadPath_big, gstrUpLoadPath_small, gstrUpLoadPath_editor
	

	Dim conn        	'数据库连接
	Dim dBType			'网站数据库类型
	Dim dBPath      	'Access数据库地址
	Dim strDBUsername, strDBPassword, strDBServerName, strDBHostIP
	
	Call OpenConn()         '连接数据库
	Call GetSiteConfig()    '获取网站配置信息
	
	'============================================
	'过程名：OpenConn
	'功  能：连接数据库
	'参  数：无
	'返回值：无
	'============================================
	Sub OpenConn()
		Dim ConnStr     '数据库连接字符串
		
		On Error Resume Next
		
		Select Case dbType
			Case DATABASE_ACCESS
				ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbpath)
			Case DATABASE_MSSQL
				ConnStr = "Driver={SQL Server};Server=" & strDBHostIP & ";Database=" & strDBServerName & ";Uid=" & strDBUsername & ";Pwd=" & strDBPassword & "; "
		End Select 
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open ConnStr
		
		If Err Then
			Response.Write Err.description
			Err.Clear
			Set conn = Nothing
			Response.Write "数据库连接出错，请检查配置文件中的数据库参数设置。"
			Response.End
		End If
	End Sub

	'============================================
	'过程名：CloseConn
	'功  能：关闭数据库连接
	'参  数：无
	'返回值：无
	'============================================
	Sub CloseConn()
		On Error Resume Next
		
		If IsObject(conn) Then
			conn.Close
			Set conn = Nothing
		End if
		
		If Err Then
			Err.Clear
		End If
	End Sub

	'============================================
	'过程名：UrlFormat
	'功  能：在地址字符串后面添加"/"
	'参  数：要处理的字符串
	'返回值：处理后的字符串
	'============================================
	Function UrlFormat(str)
		If Len(str) <= 0 Then
			str = "/"
		End If
		If Right(str,1) <> "/" then
			str = str & "/"
		End If
		UrlFormat = str
	End Function

	'============================================
	'过程名：GetSiteConfig
	'功  能：获取网站配置信息
	'参  数：无
	'返回值：无
	'============================================
	Sub GetSiteConfig()
		On Error Resume Next
		
		Dim strConfig, rsConfig
		
		gstrInstallDir = UrlFormat(gstrInstallDir)
		
		strConfig = "select * from SiteConfig"
		Set rsConfig = Server.CreateObject("ADODB.Recordset")
		rsConfig.Open strConfig, conn, 1, 3
		
		If Not (rsConfig.Bof And rsConfig.Eof) Then
			gstrSiteName = Trim(rsConfig("SiteName") & "")
			gstrSiteTitle = Trim(rsConfig("SiteTitle") & "")
			gstrSiteUrl = Trim(rsConfig("SiteUrl") & "")
			gstrInstallDir = Trim(rsConfig("InstallDir") & "")
			gstrLogoUrl = Trim(rsConfig("LogoUrl") & "")
			gstrBannerUrl = Trim(rsConfig("BannerUrl") & "")
			gstrWebmasterName = Trim(rsConfig("WebmasterName") & "")
			gstrWebmasterEmail = Trim(rsConfig("WebmasterEmail") & "")
			gstrCopyright = Trim(rsConfig("Copyright") & "")
			glngSessionTimeout = CLng(rsConfig("SessionTimeout") & "")
			glngPageSize = CLng(rsConfig("PageSize") & "")
			gstrPicFolderUrl = Trim(rsConfig("PicFolderUrl") & "")
			gstrFileFolderUrl = Trim(rsConfig("FileFolderUrl") & "")
			gstrNewsFolderUrl = Trim(rsConfig("NewsFolderUrl") & "")
			gstrAdminPanelUrl = Trim(rsConfig("AdminPanelUrl") & "")
			gstrUserPanelUrl = Trim(rsConfig("UserPanelUrl") & "")
			gstrNoPicUrl = Trim(rsConfig("NoPicUrl") & "")
			gstrFSOName = Trim(rsConfig("FSOName") & "")
			gstrNoFriendLinkPicUrl = Trim(rsConfig("NoFriendLinkPicUrl") & "")
		End If
		
		If rsConfig.State = 1 Then rsConfig.Close
		Set rsConfig = Nothing
		
		If glngSessionTimeout = 0 Then glngSessionTimeout = 20
		If glngPageSize = 0 Then glngPageSize = 20
		If Trim(gstrPicFolderUrl) = "" Then gstrPicFolderUrl = gstrInstallDir & "UpLoadFile/"
		If Trim(gstrFileFolderUrl) = "" Then gstrFileFolderUrl = gstrInstallDir & "UpLoadFile/"
		If Trim(gstrNewsFolderUrl) = "" Then gstrNewsFolderUrl = gstrInstallDir & "UpLoadFile/"
		If Trim(gstrAdminPanelUrl) = "" Then gstrAdminPanelUrl = gstrInstallDir & "abc/"
		If Trim(gstrUserPanelUrl) = "" Then gstrUserPanelUrl = gstrInstallDir & "UserPanel/"
		If Trim(gstrNoPicUrl) = "" Then gstrNoPicUrl = gstrInstallDir & "Images/nopic.gif"
		If Trim(gstrFSOName) = "" Then gstrFSOName = "Scripting.FileSystemObject"
		If Trim(gstrNoFriendLinkPicUrl) = "" Then gstrNoFriendLinkPicUrl = gstrInstallDir & "Images/friendlink.gif"
		If Trim(gstrNoFriendLink) = "" Then gstrNoFriendLink = gstrInstallDir & "Images/friendlink.gif"
		
		gstrSiteUrl = UrlFormat(gstrSiteUrl)
		gstrPicFolderUrl = UrlFormat(gstrPicFolderUrl)
		gstrFileFolderUrl = UrlFormat(gstrFileFolderUrl)
		gstrNewsFolderUrl = UrlFormat(gstrNewsFolderUrl)
		gstrAdminPanelUrl = UrlFormat(gstrAdminPanelUrl)
		gstrUserPanelUrl = UrlFormat(gstrUserPanelUrl)
		
		If gstrSessionPrefix = "" Then
			If gstrSiteUrl = "" Then
				gstrSessionPrefix = gstrInstallDir
			Else
				gstrSessionPrefix = gstrSiteUrl
			End If
		End If
		
		If Err Then
			Err.Clear
		End If
	End Sub
	
%>