<html>
<head>
<title>�ϴ��ļ�</title>
</head>
	<body style="padding:0; margin:0; font-size:12px;">
<%
	dim frm, frminput
	frm = trim(request("frm")&"")
	frminput = trim(request("frminput")&"")
%>
<script language="javascript">
function UploadData(){
	if (document.getElementById('imageurl').value == '') {
		document.getElementById('imageurl').focus();
		alert('��ѡ��Ҫ�ϴ����ļ�!');
		return false;
	}

	document.uploadfrm.action = "upload.asp?frm=<%=frm%>&frminput=<%=frminput%>&m" + Math.random();  //�����ϴ����ݵĳ���
	//document.uploadfrm.target = "upload"  
	document.uploadfrm.submit();     //�ύ��

}

function CacelUpload(){
	document.getElementById('uploadfrm').reset();
	window.returnValue = null;
	window.close();
}
</script>
<base target="_self">
<form name="uploadfrm" style="padding:0; margin:0;" id="uploadfrm" method="post" action="upload.asp?frm=<%=frm%>&frminput=<%=frminput%>&m=<%=now%>" enctype="multipart/form-data">
	<input type="file" name="imageurl" id="imageurl">
	<a onClick="UploadData()" style="background:#eee; border:1px solid #ccc; height:18px; line-height:18px; cursor:pointer;">�ϴ�</a>
	�ļ��뱣����500K����,������ܻ��ϴ�ʧ��!
</form>
</body>
</html>

