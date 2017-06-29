<!-- #include file="../Config.asp" -->
<%

	Dim dblStartTimer
	dblStartTimer = Timer()
	
	'����
	Const DATABASE_ACCESS = 1
	Const DATABASE_MSSQL = 2
	
	'ҳȫ�ֱ�����
	Dim gstrSiteName, gstrSiteTitle, gstrLogoUrl, gstrBannerUrl, gstrKeyWords, gstrdescription,gstrCompanyAbout
	Dim gstrWebmasterName, gstrWebmasterEmail, gstrCopyright
	Dim gstrSiteUrl, gstrInstallDir, glngSessionTimeout, glngPageSize, gstrNoPicUrl, gstrNoFriendLinkPicUrl, gstrNoFriendLink
	Dim gstrPicFolderUrl, gstrFileFolderUrl, gstrNewsFolderUrl, gstrAdminPanelUrl, gstrUserPanelUrl
	Dim gstrFSOName, gstrSessionPrefix
	
	Dim gstrAllowExt, gstrUpLoadPath_big, gstrUpLoadPath_small, gstrUpLoadPath_editor
	

	Dim conn        	'���ݿ�����
	Dim dBType			'��վ���ݿ�����
	Dim dBPath      	'Access���ݿ��ַ
	Dim strDBUsername, strDBPassword, strDBServerName, strDBHostIP
	
	Call OpenConn()         '�������ݿ�
	Call GetSiteConfig()    '��ȡ��վ������Ϣ
	
	'============================================
	'��������OpenConn
	'��  �ܣ��������ݿ�
	'��  ������
	'����ֵ����
	'============================================
	Sub OpenConn()
		Dim ConnStr     '���ݿ������ַ���
		
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
			Response.Write "���ݿ����ӳ������������ļ��е����ݿ�������á�"
			Response.End
		End If
	End Sub

	'============================================
	'��������CloseConn
	'��  �ܣ��ر����ݿ�����
	'��  ������
	'����ֵ����
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
	'��������UrlFormat
	'��  �ܣ��ڵ�ַ�ַ����������"/"
	'��  ����Ҫ������ַ���
	'����ֵ���������ַ���
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
	'��������GetSiteConfig
	'��  �ܣ���ȡ��վ������Ϣ
	'��  ������
	'����ֵ����
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