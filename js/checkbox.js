//实现全选
//cheageBox （触发事件的多选框名称，发生事件的多选框名称）
function cheageBox(oPar,oClt) {
	var oParState = document.all(oPar).checked;
	var oCltNAMEs = document.getElementsByName(oClt);

	for(var i=0; i<oCltNAMEs.length; i++) {
			oCltNAMEs[i].checked = oParState;
		}
	}

//把勾选的进行批量删除。
//AllDelete(要提交删除的多选框名称);
function AllDelete(oClt) {
	var oCltNAMEs = document.getElementsByName(oClt);
	var bPass = false;
	var DelConunt = 0;
	for(var i=0; i<oCltNAMEs.length; i++) {
		var oCltState = oCltNAMEs[i].checked;
		if (oCltState == true) {
			DelConunt= DelConunt+1;
			bPass = true;
		}
		
	}
	if (bPass == false) {
		alert("请选中要准备删除的记录!");
		return false;
	}
	//删除的总条数
	//alert(DelConunt);
	//提交到删除页
	if(confirm("删除记录将不能恢复,继续吗?")) {
		document.forms[0].action="alldelete.asp";
		document.forms[0].submit();
	}
}