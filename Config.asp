<%
	'=========================����վ����ǰ�������޸�������Ϣ=========================
	
	'��վ��װĿ¼������ڸ�Ŀ¼
	gstrInstallDir = "/site/book/"
	gstrKeyWords = "���Ͻ���ϵͳ"
	gstrdescription = ""
	gstrCompanyAbout = ""
	
	gstrAllowExt = "jpg|gif|rar|zip"
	gstrUpLoadPath_big = gstrInstallDir & "uppic/big/"
	gstrUpLoadPath_small = gstrInstallDir & "uppic/small/"
	gstrUpLoadPath_editor = gstrInstallDir & "uppic/editor/"
	
	'��ʱʱ�䣨��λ���֣�
	glngSessionTimeout = 20
	'����ÿҳ��ʾ��������
	glngPageSize = 20
	
	'��վ����
	gstrSiteName = "���Ͻ���ϵͳ"
	'��վ���⣨˵��)
	'gstrSiteTitle = "���Ͻ���ϵͳ"
	gstrSiteTitle = "���Ͻ���ϵͳ"
	'Session��ǰ׺
	gstrSessionPrefix = "HNAYHXYEYA_CN_"
	
	'���ݿ�����(DATABASE_ACCESS, DATABASE_MSSQL)
	dBType = DATABASE_ACCESS
	
	'ACCESS���ݿ���ļ�������ʹ���������վ��Ŀ¼�ĵľ���·��
	dBPath = gstrInstallDir & "Database/CN_HN-LXCPA.mdb"
	
	'SQL Server ���ݿ�����
	strDBUsername = "sa"                'SQL���ݿ��û���
	strDBPassword = "sa"                'SQL���ݿ��û�����
	strDBServerName = "master"         	'SQL���ݿ���
	strDBHostIP = "192.168.0.6"           'SQL����IP��ַ�����ؿ��á�127.0.0.1����(local)�����Ǳ���������ʵIP��
	
	'================================================================================
%>