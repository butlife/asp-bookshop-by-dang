function ppRoll(a)
{
	this.myA = a;
	this.myA.IsPlay = 1;
	this.$(a.demo).style.overflow = "hidden";
	this.$(a.demo).style.width = a.width;
	this.$(a.demo).style.height = a.height;
	this.$(a.demo2).innerHTML=this.$(a.demo1).innerHTML;
	this.$(a.demo).scrollTop=this.$(a.demo).scrollHeight;
	this.Marquee();
	this.$(a.demo).onmouseover=function() {eval(a.objStr+".clearIntervalpp();");}
	this.$(a.demo).onmouseout=function() {eval(a.objStr+".setTimeoutpp();")}
}
ppRoll.prototype.$ = function(Id)
{
	return document.getElementById(Id);
}
ppRoll.prototype.getV = function(){ 
alert(this.$(this.myA.demo2).offsetWidth-this.$(this.myA.demo).scrollLeft);
alert(this.$(this.myA.demo2).offsetWidth);
alert(this.$(this.myA.demo).scrollLeft);}
ppRoll.prototype.Marquee = function()
{
	this.MyMar=setTimeout(this.myA.objStr+".Marquee();",this.myA.speed);
	if(this.myA.IsPlay == 1)
	{
		//向上滚动
		if(this.myA.direction == "top")
		{
			if(this.$(this.myA.demo).scrollTop>=this.$(this.myA.demo2).offsetHeight)
				this.$(this.myA.demo).scrollTop-=this.$(this.myA.demo2).offsetHeight;
			else{
				this.$(this.myA.demo).scrollTop++;
			}
		}
		
		//向下滚动
		if(this.myA.direction == "down")
		{
			if(this.$(this.myA.demo1).offsetTop-this.$(this.myA.demo).scrollTop>=0)
				this.$(this.myA.demo).scrollTop+=this.$(this.myA.demo2).offsetHeight;
			else{
				this.$(this.myA.demo).scrollTop--;
			}
		}
		
		//向左滚动
		if(this.myA.direction == "left")
		{
			if(this.$(this.myA.demo2).offsetWidth-this.$(this.myA.demo).scrollLeft<=0)
				this.$(this.myA.demo).scrollLeft-=this.$(this.myA.demo1).offsetWidth;
			else{
				this.$(this.myA.demo).scrollLeft++;
			}
		}
		
		//向右滚动
		if(this.myA.direction == "right")
		{
			if(this.$(this.myA.demo).scrollLeft<=0)
				this.$(this.myA.demo).scrollLeft+=this.$(this.myA.demo2).offsetWidth;
			else{
				this.$(this.myA.demo).scrollLeft--;
			}
		}

	}
}
ppRoll.prototype.clearIntervalpp = function()
{
	this.myA.IsPlay = 0;
}
ppRoll.prototype.setTimeoutpp = function()
{
	this.myA.IsPlay = 1;
}
