function LTrim(s){return s.replace(/^\s*/,"");}
function RTrim(s){return s.replace(/\s*$/,"");}
function Trim(s){return RTrim(LTrim(s));}
function IsEmpty(s){var tmp_str=Trim(s);return tmp_str.length==0;}
function IsMail(s){var tmp_str=Trim(s);var pattern=/^[_a-z0-9-]+(.[_a-z0-9-]+)*@[a-z0-9-]+(.[a-z0-9-]+)*$/;return pattern.test(tmp_str);}
function IsNumber(s){var tmp_str=Trim(s);var pattern=/^[0-9]/;return pattern.test(tmp_str);}
function IsInt(s,sign,zero){var reg;var bZero;if(Trim(s)==""){return false;}
else{s=s.toString();}
if((sign==null)||(Trim(sign)=="")){sign="+-";}
if((zero==null)||(Trim(zero)=="")){bZero=false;}
else{zero=zero.toString();if(zero=="0"){bZero=true;}
else{alert("检查是否包含0参数，只可为(空、0)");}}
switch(sign){case"+-":reg=/(^-?|^\+?)\d+$/;break;case"+":if(!bZero){reg=/^\+?[0-9]*[1-9][0-9]*$/;}
else{reg=/^\+?[0-9]*[0-9][0-9]*$/;}
break;case"-":if(!bZero){reg=/^-[0-9]*[1-9][0-9]*$/;}
else{reg=/^-[0-9]*[0-9][0-9]*$/;}
break;default:alert("检查符号参数，只可为(空、+、-)");return false;break;}
var r=s.match(reg);if(r==null){return false;}
else{return true;}}
function IsFloat(s,sign,zero){var reg;var bZero;if(Trim(s)==""){return false;}
else{s=s.toString();}
if((sign==null)||(Trim(sign)=="")){sign="+-";}
if((zero==null)||(Trim(zero)=="")){bZero=false;}
else{zero=zero.toString();if(zero=="0"){bZero=true;}
else{alert("检查是否包含0参数，只可为(空、0)");}}
switch(sign){case"+-":reg=/^((-?|\+?)\d+)(\.\d+)?$/;break;case"+":if(!bZero){reg=/^\+?(([0-9]+\.[0-9]*[1-9][0-9]*)|([0-9]*[1-9][0-9]*\.[0-9]+)|([0-9]*[1-9][0-9]*))$/;}
else{reg=/^\+?\d+(\.\d+)?$/;}
break;case"-":if(!bZero){reg=/^-(([0-9]+\.[0-9]*[1-9][0-9]*)|([0-9]*[1-9][0-9]*\.[0-9]+)|([0-9]*[1-9][0-9]*))$/;}
else{reg=/^((-\d+(\.\d+)?)|(0+(\.0+)?))$/;}
break;default:alert("检查符号参数，只可为(空、+、-)");return false;break;}
var r=s.match(reg);if(r==null){return false;}
else{return true;}}
function IsEnLetter(s,size){var reg;if(Trim(s)==""){return false;}
else{s=s.toString();}
if((size==null)||(Trim(size)=="")){size="UL";}
else{size=size.toUpperCase();}
switch(size){case"UL":reg=/^[A-Za-z]+$/;break;case"U":reg=/^[A-Z]+$/;break;case"L":reg=/^[a-z]+$/;break;default:alert("检查大小写参数，只可为(空、UL、U、L)");return false;break;}
var r=s.match(reg);if(r==null){return false;}
else{return true;}}
function IsColor(color){var temp=color;if(temp=="")return true;if(temp.length!=7)return false;return(temp.search(/\#[a-fA-F0-9]{6}/)!=-1);}
function IsURL(url){var sTemp;var b=true;sTemp=url.substring(0,7);sTemp=sTemp.toUpperCase();if((sTemp!="HTTP://")||(url.length<10)){b=false;}
return b;}
function IsMobile(s){var tmp_str=Trim(s);var pattern=/13\d{9}/;return pattern.test(tmp_str);}