//ʵ��ȫѡ
//cheageBox �������¼��Ķ�ѡ�����ƣ������¼��Ķ�ѡ�����ƣ�
function cheageBox(oPar,oClt) {
	var oParState = document.all(oPar).checked;
	var oCltNAMEs = document.getElementsByName(oClt);

	for(var i=0; i<oCltNAMEs.length; i++) {
			oCltNAMEs[i].checked = oParState;
		}
	}

//�ѹ�ѡ�Ľ�������ɾ����
//AllDelete(Ҫ�ύɾ���Ķ�ѡ������);
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
		alert("��ѡ��Ҫ׼��ɾ���ļ�¼!");
		return false;
	}
	//ɾ����������
	//alert(DelConunt);
	//�ύ��ɾ��ҳ
	if(confirm("ɾ����¼�����ָܻ�,������?")) {
		document.forms[0].action="alldelete.asp";
		document.forms[0].submit();
	}
}