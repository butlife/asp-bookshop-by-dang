function ppRoll4(a){
	this.myA = a;
	
	if(a.direction == "top" || a.direction == "down") this.writeTopDown();
	if(a.direction == "left" || a.direction == "right") this.writeLeftRight();

	this.myA.IsPlay = 1;
	this.$(a.objStr+"demo").style.overflow = "hidden";
	this.$(a.objStr+"demo").style.width = a.width;
	this.$(a.objStr+"demo").style.height = a.height;
	this.$(a.objStr+"demo2").innerHTML=this.$(a.objStr+"demo1").innerHTML;
	this.$(a.objStr+"demo3").innerHTML=this.$(a.objStr+"demo1").innerHTML;
	this.$(a.objStr+"demo").scrollTop=this.$(a.objStr+"demo").scrollHeight;
	this.Marquee();
	this.$(a.objStr+"demo").onmouseover=function() {eval(a.objStr+".clearIntervalpp();");}
	this.$(a.objStr+"demo").onmouseout=function() {eval(a.objStr+".setTimeoutpp();")}
	
	
}
ppRoll4.prototype.writeTopDown =function()
{
	document.write("<div id=\""+this.myA.objStr+"demo\"><div id=\""+this.myA.objStr+"demo1\">");
				
	//document.write("<img src='"+this.myA.picSrcArr[0].picSrc+"' />");
	//document.write(this.myA.picSrcArr.length);
	document.write("<table width=\""+this.myA.tdWidth+"\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">");
	for(var i in this.myA.picSrcArr) 
	{
		document.write("<tr><td align=\"center\" valign=\"middle\" width='"+this.myA.tdWidth+"' height='"+this.myA.tdHeight+"'>");
		document.write(this.myA.picTemplate.replace(new RegExp("\\$picSrc","g"),this.myA.picSrcArr[i].picSrc).replace(new RegExp("\\$picHref","g"),this.myA.picSrcArr[i].picHref).replace(new RegExp("\\$picTitle","g"),this.myA.picSrcArr[i].picTitle));
		//document.write("<img src='"+this.myA.picSrcArr[i].picSrc+"' />");
		document.write("</td></tr>");
	}
	document.write("</table>");
	//document.write(parseInt(this.myA.picWidth.replace("px",""))+parseInt(this.myA.picBorder.replace("px","")));
	//document.write(this.myA.tdWidth+" "+this.myA.tdHeight);
	
	document.write("</div><div id=\""+this.myA.objStr+"demo2\"></div><div id=\""+this.myA.objStr+"demo3\"></div></div>");
	
}

ppRoll4.prototype.writeLeftRight =function()
{
	document.write("<div id=\""+this.myA.objStr+"demo\"><table border=0 cellspacing=0 cellpadding=0><tr><td id=\""+this.myA.objStr+"demo1\">");
				
	//document.write("<img src='"+this.myA.picSrcArr[0].picSrc+"' />");
	//document.write(this.myA.picSrcArr.length);
	document.write("<table width=\""+parseInt(this.myA.tdWidth.replace("px","")) * parseInt(this.myA.picSrcArr.length) +"\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\"><tr>");
	for(var i in this.myA.picSrcArr) 
	{
		document.write("<td align=\"center\" valign=\"middle\" width='"+this.myA.tdWidth+"' height='"+this.myA.tdHeight+"'>");
		document.write(this.myA.picTemplate.replace(new RegExp("\\$picSrc","g"),this.myA.picSrcArr[i].picSrc).replace(new RegExp("\\$picHref","g"),this.myA.picSrcArr[i].picHref).replace(new RegExp("\\$picTitle","g"),this.myA.picSrcArr[i].picTitle));
		//document.write("<img src='"+this.myA.picSrcArr[i].picSrc+"' />");
		document.write("</td>");
	}
	document.write("</tr></table>");
	//document.write(parseInt(this.myA.picWidth.replace("px",""))+parseInt(this.myA.picBorder.replace("px","")));
	//document.write(this.myA.tdWidth+" "+this.myA.tdHeight);
	
	document.write("</td><td id=\""+this.myA.objStr+"demo2\"></td><td id=\""+this.myA.objStr+"demo3\"></td></tr></table>");
	
}

ppRoll4.prototype.$ = function(Id)
{
	return document.getElementById(Id);
}
ppRoll4.prototype.Marquee = function()
{
	this.MyMar=setTimeout(this.myA.objStr+".Marquee();",this.myA.speed);
	if(this.myA.IsPlay == 1)
	{
		//向上滚动
		if(this.myA.direction == "top")
		{
			if(this.$(this.myA.objStr+"demo").scrollTop>=this.$(this.myA.objStr+"demo2").offsetHeight)
				this.$(this.myA.objStr+"demo").scrollTop-=this.$(this.myA.objStr+"demo2").offsetHeight;
			else{
				this.$(this.myA.objStr+"demo").scrollTop++;
			}
		}
		
		//向下滚动
		if(this.myA.direction == "down")
		{
			if(this.$(this.myA.objStr+"demo1").offsetTop-this.$(this.myA.objStr+"demo").scrollTop>=0)
				this.$(this.myA.objStr+"demo").scrollTop+=this.$(this.myA.objStr+"demo2").offsetHeight;
			else{
				this.$(this.myA.objStr+"demo").scrollTop--;
			}
		}
		
		//向左滚动
		if(this.myA.direction == "left")
		{
			if(this.$(this.myA.objStr+"demo2").offsetWidth-this.$(this.myA.objStr+"demo").scrollLeft<=0)
				this.$(this.myA.objStr+"demo").scrollLeft-=this.$(this.myA.objStr+"demo1").offsetWidth;
			else{
				this.$(this.myA.objStr+"demo").scrollLeft++;
			}
		}
		
		//向右滚动
		if(this.myA.direction == "right")
		{
			if(this.$(this.myA.objStr+"demo").scrollLeft<=0)
				this.$(this.myA.objStr+"demo").scrollLeft+=this.$(this.myA.objStr+"demo2").offsetWidth;
			else{
				this.$(this.myA.objStr+"demo").scrollLeft--;
			}
		}

	}
}
ppRoll4.prototype.clearIntervalpp = function()
{
	this.myA.IsPlay = 0;
}
ppRoll4.prototype.setTimeoutpp = function()
{
	this.myA.IsPlay = 1;
}