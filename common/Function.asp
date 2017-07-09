<%
	Function CleanHTML(str)
		'����HTML
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
	'��������ReplaceBadChar
	'��  �ã����˷Ƿ���SQL�ַ�
	'��  ����expression-----Ҫ���˵��ַ�
	'����ֵ�����˺���ַ�
	'*============================================
	Function ReplaceBadChar(expression)
		If expression = "" Then
			ReplaceBadChar = ""
		Else
			ReplaceBadChar = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(expression, "'", ""), "*", ""), "?", ""), "(", ""), ")", ""), "<", ""),"-",""), ".", "")
		End If
	End Function

	'============================================
	'��������IsBoolean
	'��  �ã��ж��ַ��Ƿ�Ϊ�߼���
	'��  ����expression-----Ҫ�жϵ��ַ�
	'����ֵ���Ƿ�
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
	' ���ַ�������HTML����,�滻Server.HTMLEncode
	' ȥ��Html��ʽ��������ʾ���
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
	' ȥ��Html��ʽ�����ڴ����ݿ���ȡ��ֵ���������ʱ
	' ע�⣺value="?"���һ��Ҫ��˫����
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
	'��������IsObjInstalled
	'��  �ã��������Ƿ��Ѿ���װ
	'��  ����strClassString ----�����
	'����ֵ��True  ----�Ѿ���װ
	'       False ----û�а�װ
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
	'��������LengthStr
	'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
	'��  ����str  ----Ҫ�󳤶ȵ��ַ���
	'����ֵ���ַ�������
	' ============================================
	Function LengthStr(str)
	
		On Error Resume Next
		
		Dim WINNT_CHINESE
		
		WINNT_CHINESE = (Len("�й�") = 2)
		
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
	'������FoundInArr
	'��  �ã����һ������������Ԫ���Ƿ����ָ���ַ���
	'��  ����strArr     ----�洢�������ݵ��ִ�
	'       strToFind    ----Ҫ���ҵ��ַ���
	'       strSplit    ----����ķָ���
	'����ֵ��True,False
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
	'��������IsValidEmail
	'��  �ã����Email��ַ�Ϸ���
	'��  ����email ----Ҫ����Email��ַ
	'����ֵ��True  ----Email��ַ�Ϸ�
	'       False ----Email��ַ���Ϸ�
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
	'��������LPad
	'��  �ã����ַ���ǰ������ַ�
	'��  ����str ----Ҫ�����ַ���
	'        imax -- �����ַ�������
	'        char -- Ҫ�����ַ�
	'����ֵ��������ֵ
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
	'��������RPad
	'��  �ã����ַ����������ַ�
	'��  ����str ----Ҫ�����ַ���
	'        imax -- �����ַ�������
	'        char -- Ҫ�����ַ�
	'����ֵ��������ֵ
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
	'��������ConvertInt
	'��  �ã�ת��Ϊ����
	'��  ����str ----Ҫת�����ַ���
	'����ֵ��������ֵ
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
	'��������ConvertLong
	'��  �ã�ת��Ϊ������
	'��  ����str ----Ҫת�����ַ���
	'����ֵ����������ֵ
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
	'��������ConvertDouble
	'��  �ã�ת��Ϊ˫������
	'��  ����str ----Ҫת�����ַ���
	'����ֵ��˫��������ֵ
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
	'��������ConvertBoolean
	'��  �ã�ת��Ϊ������
	'��  ����str ----Ҫת�����ַ���
	'����ֵ����������ֵ
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
	'��������TdString
	'��  �ã�Ϊ���Html, ��ʽ���ַ���
	'��  ����str   ��Ϣ
	'����ֵ���ַ���
	'============================================
	Function TdString(str)
		If Trim(str) <> "" Then
			TdString = str
		Else
			TdString = "&nbsp;"
		End If
	End Function

	'============================================
	'��������TdDate
	'��  �ã�Ϊ���Html, ��ʽ������
	'��  ����str   ��Ϣ
	'����ֵ���ַ���
	'============================================
	Function TdDate(str)
		If Trim(str) <> "" Then
			TdDate = FormatDateTime(str, vbShortDate)
		Else
			TdDate = "&nbsp;"
		End If
	End Function

	'============================================
	'��������TdNumeric
	'��  �ã�Ϊ���Html, ��ʽ������
	'��  ����str   ��Ϣ
	'����ֵ���ַ���
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
	'��������RemoveHTML
	'��  �ã�����HTML
	'��  ����str   ��Ϣ
	'����ֵ���ַ���
	'============================================
Function RemoveHTML(strHTML) 
	Dim objRegExp, Match, Matches 
	Set objRegExp = New Regexp 
	objRegExp.IgnoreCase = True 
	objRegExp.Global = True 
	'ȡ�պϵ�<> 
	objRegExp.Pattern = "<.+?>" 
	'����ƥ�� 
	Set Matches = objRegExp.Execute(strHTML) 
	' ����ƥ�伯�ϣ����滻��ƥ�����Ŀ 
	For Each Match in Matches 
	strHtml=Replace(strHTML,Match.Value,"") 
	Next 
	RemoveHTML=strHTML 
	Set objRegExp = Nothing 
End Function
'	============================================
	' ��ʽ��ʱ��(��ʾ)
	' ������n_Flag
	' 1:"yyyy-mm-dd hh:mm:ss"
	' 2:"yyyy-mm-dd"
	' 3:"hh:mm:ss"
	' 4:"yyyy��mm��dd��"
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
	  ' yyyy��mm��dd��
	  Format_Time = y & "��" & m & "��" & d & "��"
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
	 End Select
	End Function


	function Cnum(pNum)
		select case pNum
		case 1:Cnum="һ"
		case 2:Cnum="��"
		case 3:Cnum="��"
		case 4:Cnum="��"
		case 5:Cnum="��"
		case 6:Cnum="��"
		case 7:Cnum="��"
		case 8:Cnum="��"
		case 9:Cnum="��"
		case 0:Cnum="��"
		end select
	end function
	
	
	function DBC2SBC(str,flag)
	'ȫ���ת�� flag = 0Ϊ ��תȫ��=1Ϊȫת��
	 dim i, str1
	 if len(str)<=0 then 
	 REsponse.Write("�ַ�������Ϊ��") 
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
		str = replace(str,",","��")
		str = replace(str,".","��")
		str = replace(str,"~","��")
		str = replace(str,"!","��")
		str = replace(str,"@","��")
		str = replace(str,"#","��")
		str = replace(str,"$","��")
		str = replace(str,"%","��")
		str = replace(str,"^","����")
		str = replace(str,"&","��")
		str = replace(str,"*","��")
		str = replace(str,"(","��")
		str = replace(str,")","��")
		str = replace(str,"-","��")
		str = replace(str,"=","��")
		str = replace(str,"+","��")
		str = replace(str,"\","��")
		str = replace(str,"[","��")
		str = replace(str,"]","��")
		str = replace(str,";","��")
		str = replace(str,"'","��")
		str = replace(str,":","��")
		str = replace(str,"?","��")
		str = replace(str,"""","��")
		str = replace(str,"/","��")
		str = replace(str,"<","��")
		str = replace(str,">","��")
		replaceCode = str
	end function
%>

