<%
	'=========================在网站运行前，请先修改以下信息=========================
	
	'网站安装目录，相对于根目录
	gstrInstallDir = "/"
	gstrKeyWords = "网上借书系统"
	gstrdescription = ""
	gstrServiceTel = "0372110"
	
	gstrAllowExt = "jpg|gif|rar|zip"
	gstrUpLoadPath_big = gstrInstallDir & "uppic/big/"
	gstrUpLoadPath_small = gstrInstallDir & "uppic/small/"
	gstrUpLoadPath_editor = gstrInstallDir & "uppic/editor/"
	
	'超时时间（单位：分）
	glngSessionTimeout = 20
	'网格每页显示数据行数
	glngPageSize = 20
	glngPageSize_phone = 10
	
	'网站名称
	gstrSiteName = "网上借书系统"
	'网站标题（说明)
	'gstrSiteTitle = "网上借书系统"
	gstrSiteTitle = "网上借书系统"
	'Session的前缀
	gstrSessionPrefix = "HNAYHXYEYA_CN_"
	
	'数据库类型(DATABASE_ACCESS, DATABASE_MSSQL)
	dBType = DATABASE_ACCESS
	
	'ACCESS数据库的文件名，请使用相对于网站根目录的的绝对路径
	dBPath = gstrInstallDir & "Database/CN_HN-LXCPA.mdb"
	
	'SQL Server 数据库配置
	strDBUsername = "sa"                'SQL数据库用户名
	strDBPassword = "sa"                'SQL数据库用户密码
	strDBServerName = "master"         	'SQL数据库名
	strDBHostIP = "192.168.0.6"           'SQL主机IP地址（本地可用“127.0.0.1”或“(local)”，非本机请用真实IP）
	
	'================================================================================
%>