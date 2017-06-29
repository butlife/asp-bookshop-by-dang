<html>
<head>
<title>上传文件</title>
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
		alert('请选择要上传的文件!');
		return false;
	}

	document.uploadfrm.action = "upload.asp?frm=<%=frm%>&frminput=<%=frminput%>&m" + Math.random();  //处理上传数据的程序
	//document.uploadfrm.target = "upload"  
	document.uploadfrm.submit();     //提交表单

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
	<a onClick="UploadData()" style="background:#eee; border:1px solid #ccc; height:18px; line-height:18px; cursor:pointer;">上传</a>
	文件请保持在500K以内,否则可能会上传失败!
</form>
</body>
</html>

