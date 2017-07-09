<%
	Function CleanHTML(str)
		'过滤HTML
		Dim re, strContent
		Set re=new RegExp 
		re.IgnoreCase =True 
		re.Global=True 
		re.Pattern="<(.[^>]*)>" 
		strContent = re.Replace(str,"")
		CleanHTML = Replace(strContent, "&nbsp;", "")
		set re=Nothing 
	End Function 

	'============================================
	'函数名：ReplaceBadChar
	'作  用：过滤非法的SQL字符
	'参  数：expression-----要过滤的字符
	'返回值：过滤后的字符
	'*============================================
	Function ReplaceBadChar(expression)
		If expression = "" Then
			ReplaceBadChar = ""
		Else
			ReplaceBadChar = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(expression, "'", ""), "*", ""), "?", ""), "(", ""), ")", ""), "<", ""),"-",""), ".", "")
		End If
	End Function

	'============================================
	'函数名：IsBoolean
	'作  用：判断字符是否为逻辑型
	'参  数：expression-----要判断的字符
	'返回值：是否
	'*============================================
	Function IsBoolean(expression)
		IsBoolean = False
		
		expression = Trim(expression)
		
		If expression <> "" Then
			If expression = "True" Or expression = "False" Then
				IsBoolean = True
			End If
			If expression = "0" Or expression = "1" Then
				IsBoolean = True
			End If
		End If
	End Function
	
	' ============================================
	' 把字符串进行HTML解码,替换Server.HTMLEncode
	' 去除Html格式，用于显示输出
	' ============================================
	Function outHTML(str)
		Dim sTemp
		sTemp = str
		outHTML = ""
		If IsNull(sTemp) = True Then
			Exit Function
		End If
		sTemp = Replace(sTemp, "&amp;", "&")
		sTemp = Replace(sTemp, "&lt;", "<")
		sTemp = Replace(sTemp, "&gt;", ">")
		sTemp = Replace(sTemp, "&quot;", Chr(34))
		sTemp = Replace(sTemp, "<br>", Chr(10))
		outHTML = sTemp
	End Function
	
	' ============================================
	' 去除Html格式，用于从数据库中取出值填入输入框时
	' 注意：value="?"这边一定要用双引号
	' ============================================
	Function inHTML(str)
		Dim sTemp
		sTemp = str
		inHTML = ""
		If IsNull(sTemp) = True Then
			Exit Function
		End If
		sTemp = Replace(sTemp, "&", "&amp;")
		sTemp = Replace(sTemp, "<", "&lt;")
		sTemp = Replace(sTemp, ">", "&gt;")
		sTemp = Replace(sTemp, Chr(34), "&quot;")
		inHTML = sTemp
	End Function

	' ============================================
	'函数名：IsObjInstalled
	'作  用：检查组件是否已经安装
	'参  数：strClassString ----组件名
	'返回值：True  ----已经安装
	'       False ----没有安装
	' ============================================
	Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
	
	' ============================================
	'函数名：LengthStr
	'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
	'参  数：str  ----要求长度的字符串
	'返回值：字符串长度
	' ============================================
	Function LengthStr(str)
	
		On Error Resume Next
		
		Dim WINNT_CHINESE
		
		WINNT_CHINESE = (Len("中国") = 2)
		
		If WINNT_CHINESE Then
			Dim l, t, c
			Dim i
			l = Len(str)
			t = l
			For i = 1 To l
				c = Asc(Mid(str, i, 1))
				If c < 0 Then c = c + 65536
				If c > 255 Then
					t = t + 1
				End If
			Next
			LengthStr = t
		Else
			LengthStr = Len(str)
		End If
		If Err.Number <> 0 Then Err.Clear
		
	End Function
	
	' ============================================
	'函数：FoundInArr
	'作  用：检查一个数组中所有元素是否包含指定字符串
	'参  数：strArr     ----存储数据数据的字串
	'       strToFind    ----要查找的字符串
	'       strSplit    ----数组的分隔符
	'返回值：True,False
	' ============================================
	Public Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True
				Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then
			FoundInArr = True
			End If
		End If
	End Function
	
	'============================================
	'函数名：IsValidEmail
	'作  用：检查Email地址合法性
	'参  数：email ----要检查的Email地址
	'返回值：True  ----Email地址合法
	'       False ----Email地址不合法
	'============================================
	Function IsValidEmail(email)
		Dim names, name, i, c
		IsValidEmail = True
		names = Split(email, "@")
		If UBound(names) <> 1 Then
		   IsValidEmail = False
		   Exit Function
		End If
		For Each name In names
			If Len(name) <= 0 Then
			IsValidEmail = False
			Exit Function
			End If
			For i = 1 To Len(name)
			c = LCase(Mid(name, i, 1))
			If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
			   IsValidEmail = False
			   Exit Function
			 End If
		   Next
		   If Left(name, 1) = "." Or Right(name, 1) = "." Then
			  IsValidEmail = False
			  Exit Function
		   End If
		Next
		If InStr(names(1), ".") <= 0 Then
			IsValidEmail = False
		   Exit Function
		End If
		i = Len(names(1)) - InStrRev(names(1), ".")
		If i <> 2 And i <> 3 And i <> 4 Then
		   IsValidEmail = False
		   Exit Function
		End If
		If InStr(email, "..") > 0 Then
		   IsValidEmail = False
		End If
	End Function

	'============================================
	'函数名：LPad
	'作  用：在字符串前边填充字符
	'参  数：str ----要填充的字符串
	'        imax -- 填充后字符串长度
	'        char -- 要填充的字符
	'返回值：整型数值
	'============================================
	Function LPad(str, imax, char)
		Dim ilen
		ilen = Len(str)
		If ilen >= imax Then
			LPad = str
			Exit Function
		End If
		
		LPad = String(imax - ilen, char) & str
		
	End Function
	
	'============================================
	'函数名：RPad
	'作  用：在字符串后边填充字符
	'参  数：str ----要填充的字符串
	'        imax -- 填充后字符串长度
	'        char -- 要填充的字符
	'返回值：整型数值
	'============================================
	Function RPad(str, imax, char)
		Dim ilen
		ilen = Len(str)
		If ilen >= imax Then
			RPad = str
			Exit Function
		End If
		
		RPad = str & String(imax - ilen, char)
		
	End Function
	
	'============================================
	'函数名：ConvertInt
	'作  用：转换为整型
	'参  数：str ----要转换的字符串
	'返回值：整型数值
	'============================================
	Function ConvertInt(str)
		ConvertInt = ConvertIntForDefault(str, 0)
	End Function
	
	Function ConvertIntForDefault(str, default)
		If IsNumeric(default) Then 
			default = CInt(default)
		Else
			default = 0
		End If
		
		If IsNumeric(str) Then
			ConvertIntForDefault = CInt(str)
		Else
			ConvertIntForDefault = default
		End If
	End Function
	
	'============================================
	'函数名：ConvertLong
	'作  用：转换为长整型
	'参  数：str ----要转换的字符串
	'返回值：长整型数值
	'============================================
	Function ConvertLong(str)
		ConvertLong = ConvertLongForDefault(str, 0)
	End Function
	
	Function ConvertLongForDefault(str, default)
		If IsNumeric(default) Then 
			default = CLng(default)
		Else
			default = 0
		End If
		
		If IsNumeric(str) Then
			ConvertLongForDefault = CLng(str)
		Else
			ConvertLongForDefault = default
		End If
	End Function
	
	'============================================
	'函数名：ConvertDouble
	'作  用：转换为双精度型
	'参  数：str ----要转换的字符串
	'返回值：双精度型数值
	'============================================
	Function ConvertDouble(str)
		ConvertDouble = ConvertDoubleForDefault(str, 0)
	End Function
	
	Function ConvertDoubleForDefault(str, default)
		If IsNumeric(default) Then 
			default = CDbl(default)
		Else
			default = 0
		End If
		
		If IsNumeric(str) Then
			ConvertDoubleForDefault = CDbl(str)
		Else
			ConvertDoubleForDefault = default
		End If
	End Function
	
	'============================================
	'函数名：ConvertBoolean
	'作  用：转换为布尔型
	'参  数：str ----要转换的字符串
	'返回值：布尔型数值
	'============================================
	Function ConvertBoolean(str)
		ConvertBoolean = ConvertBooleanForDefault(str, False)
	End Function
	
	Function ConvertBooleanForDefault(str, default)
		If IsBoolean(default) Then 
			default = CBool(default)
		Else
			default = False
		End If
		
		If IsBoolean(str) Then
			ConvertBooleanForDefault = CBool(str)
		Else
			ConvertBooleanForDefault = default
		End If
	End Function
	
	'============================================
	'过程名：TdString
	'作  用：为输出Html, 格式化字符串
	'参  数：str   信息
	'返回值：字符串
	'============================================
	Function TdString(str)
		If Trim(str) <> "" Then
			TdString = str
		Else
			TdString = "&nbsp;"
		End If
	End Function

	'============================================
	'过程名：TdDate
	'作  用：为输出Html, 格式化日期
	'参  数：str   信息
	'返回值：字符串
	'============================================
	Function TdDate(str)
		If Trim(str) <> "" Then
			TdDate = FormatDateTime(str, vbShortDate)
		Else
			TdDate = "&nbsp;"
		End If
	End Function

	'============================================
	'过程名：TdNumeric
	'作  用：为输出Html, 格式化数字
	'参  数：str   信息
	'返回值：字符串
	'============================================
	Function TdNumeric(str)
		If Trim(str) <> "" Then
			If CDbl(str) <> 0 Then
				TdNumeric = Trim(str)
			Else
				TdNumeric = "&nbsp;"
			End If
		Else
			TdNumeric = "&nbsp;"
		End If
	End Function

	'============================================
	'过程名：RemoveHTML
	'作  用：过滤HTML
	'参  数：str   信息
	'返回值：字符串
	'============================================
Function RemoveHTML(strHTML) 
	Dim objRegExp, Match, Matches 
	Set objRegExp = New Regexp 
	objRegExp.IgnoreCase = True 
	objRegExp.Global = True 
	'取闭合的<> 
	objRegExp.Pattern = "<.+?>" 
	'进行匹配 
	Set Matches = objRegExp.Execute(strHTML) 
	' 遍历匹配集合，并替换掉匹配的项目 
	For Each Match in Matches 
	strHtml=Replace(strHTML,Match.Value,"") 
	Next 
	RemoveHTML=strHTML 
	Set objRegExp = Nothing 
End Function
'	============================================
	' 格式化时间(显示)
	' 参数：n_Flag
	' 1:"yyyy-mm-dd hh:mm:ss"
	' 2:"yyyy-mm-dd"
	' 3:"hh:mm:ss"
	' 4:"yyyy年mm月dd日"
	' 5:"yyyymmdd"
	' ============================================
	Function Format_Time(s_Time, n_Flag)
	 Dim y, m, d, h, mi, s
	 Format_Time = ""
	 If IsDate(s_Time) = False Then Exit Function
	 y = cstr(year(s_Time))
	 m = cstr(month(s_Time))
	 If len(m) = 1 Then m = "0" & m
	 d = cstr(day(s_Time))
	 If len(d) = 1 Then d = "0" & d
	 h = cstr(hour(s_Time))
	 If len(h) = 1 Then h = "0" & h
	 mi = cstr(minute(s_Time))
	 If len(mi) = 1 Then mi = "0" & mi
	 s = cstr(second(s_Time))
	 If len(s) = 1 Then s = "0" & s
	 Select Case n_Flag
	 Case 1
	  ' yyyy-mm-dd hh:mm:ss
	  Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	 Case 2
	  ' yyyy-mm-dd
	  Format_Time = y & "-" & m & "-" & d
	 Case 3
	  ' hh:mm:ss
	  Format_Time = h & ":" & mi & ":" & s
	 Case 4
	  ' yyyy年mm月dd日
	  Format_Time = y & "年" & m & "月" & d & "日"
	 Case 5
	  ' yyyymmdd
	  Format_Time = y & m & d
	 Case 6
	  y = right(y, 2)
	  Format_Time = y & "/" & m & "/" & d & " " & h & ":" & mi
	 Case 7
	  Format_Time = m & "/" & d & " " & h & ":" & mi
	 Case 8
	  y = right(y, 2)
	  Format_Time = y & "/" & m & "/" & d
	 Case 9
	  Format_Time = "(" & m & "/" & d & ")"
	 Case 10
	  Format_Time = "(" & m & d & "&nbsp;" & h & ":" & mi & ")"
	 Case 11
	  Format_Time = y & "/" & m & "/" & d
	 End Select
	End Function


	function Cnum(pNum)
		select case pNum
		case 1:Cnum="一"
		case 2:Cnum="二"
		case 3:Cnum="三"
		case 4:Cnum="四"
		case 5:Cnum="五"
		case 6:Cnum="六"
		case 7:Cnum="七"
		case 8:Cnum="八"
		case 9:Cnum="九"
		case 0:Cnum="零"
		end select
	end function
	
	function DBC2SBC(str,flag)
	'全半角转换 flag = 0为 半转全，=1为全转半
	 dim i
	 if len(str)<=0 then 
	 REsponse.Write("字符串参数为空") 
	exit function 
	 end if
	 for i=1 to len(str)
	 str1=asc(mid(str,i,1))
	 if str1>0 and str1<=125 and not flag then
	 dbc2sbc=dbc2sbc&chr(asc(mid(str,i,1))-23680)
	 else
	 dbc2sbc=dbc2sbc&chr(asc(mid(str,i,1))+23680)
	 end if
	 next
	End function
	
	function replaceCode(str)
		str = replace(str,",","，")
		str = replace(str,".","。")
		str = replace(str,"~","～")
		str = replace(str,"!","！")
		str = replace(str,"@","＠")
		str = replace(str,"#","＃")
		str = replace(str,"$","￥")
		str = replace(str,"%","％")
		str = replace(str,"^","……")
		str = replace(str,"&","＆")
		str = replace(str,"*","×")
		str = replace(str,"(","（")
		str = replace(str,")","）")
		str = replace(str,"-","－")
		str = replace(str,"=","＝")
		str = replace(str,"+","＋")
		str = replace(str,"\","＼")
		str = replace(str,"[","【")
		str = replace(str,"]","】")
		str = replace(str,";","；")
		str = replace(str,"'","‘")
		str = replace(str,":","：")
		str = replace(str,"?","？")
		str = replace(str,"""","“")
		str = replace(str,"/","、")
		str = replace(str,"<","《")
		str = replace(str,">","》")
		replaceCode = str
	end function
%>

