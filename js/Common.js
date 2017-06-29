window.onerror=function(){return false};

function openWindow(url,iWidth,iHeight) {
	var iLeft = (screen.availWidth - iWidth)/2;
	var iTop = (screen.availHeight - iHeight)/2;
	
	window.open(url, "_blank", "toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=no,resizable=no,width=" + iWidth + ",height=" + iHeight + ",top= " + iTop + ",left=" + iLeft + "");

	/*window.location.href = url;*/

}

function openDialog(url, iWidth, iHeigth){
	var sURL = url;
	var vArguments = "";
	var sFeatures = "dialogHeight: " + iHeigth + "px; dialogWidth: " + iWidth + "px; location:no; edge: Raised; center: Yes; help: No; resizable: No; status: No;";
	
	window.showModalDialog(url, vArguments, sFeatures);
}

function PrintForm() {
	alert ("´òÓ¡ÏµÍ³!");
}

function CheckInput(str, iType) {
	var re;
	switch (iType) {
		case 1:
			re = /[^\d]/g;
			break;
		case 2:
			re = /[^\w+$]/g;
			break;
		/*case 3:
			re = /^(d{2}|d{4})-((0([1-9]{1}))|(1[1|2]))-(([0-2]([1-9]{1}))|(3[0|1]))$/g;
			//    (\d{4}|\d{2})((0([1-9]{1}))|(1([1|2]{1})))((([0-2]{1})([1-9]{1}))|(3([0|1]{1})))$
			break;*/
	 }
	
	if(re.test(str))
		return true;
	else
		return false;
}

function DrawImage(ImgD,iwidth,iheight){
	var flag=false;
    var image=new Image();
    image.src=ImgD.src;
    if(image.width>0 && image.height>0){
    flag=true;
    if(image.width/image.height>= iwidth/iheight){
        if(image.width>iwidth){
        ImgD.width=iwidth;
        ImgD.height=(image.height*iwidth)/image.width;
        }else{
        ImgD.width=image.width;
        ImgD.height=image.height;
        }
        }
    else{
        if(image.height>iheight){
        ImgD.height=iheight;
        ImgD.width=(image.width*iheight)/image.height;
        }else{
        ImgD.width=image.width;
        ImgD.height=image.height;
        }
        }
    }
}
function picload() {
 var iwidth = 480;
 var iheight = 600;
 var mainDiv = document.getElementById("WordPic_contact");
 if (!mainDiv) {return;}
 var oIMG = mainDiv.getElementsByTagName("IMG");
 if (oIMG.length>0) {
 for (var i=0;i<oIMG.length; i++) {
  var image = new Image();
  var ImgD = oIMG[i]
  image.src = ImgD.src;
  if(image.width>0 && image.height>0){
   if(image.width/image.height>= iwidth/iheight){if(image.width>iwidth){ImgD.width=iwidth;ImgD.height=(image.height*iwidth)/image.width;}else{ImgD.width=image.width;ImgD.height=image.height;}}
   else{if(image.height>iheight){ImgD.height=iheight;ImgD.width=(image.width*iheight)/image.height;}else{ImgD.width=image.width;ImgD.height=image.height;}}
  }
 }
}
}

