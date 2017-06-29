<style type="text/css">
<!--
body{
	background:#337abb;
	margin:0px;
	padding:0px;
	text-decoration: none;
}

.menuBox{
    width:3px;
	height:100%;
	vertical-align:middle;
}
-->
</style>
<script language="javascript" type="text/javascript">
function switchMenuBar(str){
	src = event.srcElement;
	if(str != "Images/arrow_r.gif"){
		if(src.expand == true){
			parent.frameLeft.cols = "180,10,*";
			document.all.timg.src = "Images/arrow_l.gif";
			document.all.timg.title="Òþ²Ø×óÀ¸";
			src.expand = false;
		}
		else{
			parent.frameLeft.cols = "0,10,*";
			document.all.timg.src = "Images/arrow_r.gif";
			document.all.timg.title="Õ¹¿ª×óÀ¸";
			src.expand = true;
		}
	}
}
</script>
<table height="100%" border="0" cellpadding="0" cellspacing="0" class="menuBox">
  <tr>
    <td height="40%"></td>
  <tr>
    <td height="10%" align="left" valign="middle"><Img src="Images/arrow_l.gif" name="timg" hspace="0" vspace="0" border='0' align="left" id="timg" style="cursor:pointer;margin:1px;padding:0;" title="Òþ²Ø×óÀ¸" onclick="switchMenuBar(this.src);" /></td>
  <tr>
    <td height="50%"></td>
  </tr>
</table>
