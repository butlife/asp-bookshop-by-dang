<%
Dim DoteyUpload_SourceData
Class DoteyUpload
	
	Public Files
	Public Form
	Public MaxTotalBytes
	Public Version
	Public ProgressID
	Public ErrMsg
	
	Private BytesRead
	Private ChunkReadSize
	Private Info
	Private Progress

	Private UploadProgressInfo
	Private CrLf

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")	' 上传文件集合
		Set Form = Server.CreateObject("Scripting.Dictionary")	' 表单集合
'		Set Files = new DictionaryClass	' 上传文件集合
'		Set Form = new DictionaryClass	' 表单集合
		UploadProgressInfo = "DoteyUploadProgressInfo"  ' Application的Key
		MaxTotalBytes = 1 *1024 *1024 *1024 ' 默认最大1G
		ChunkReadSize = 64 * 1024	' 分块大小64K
		CrLf = Chr(13) & Chr(10)	' 换行

		Set DoteyUpload_SourceData = Server.CreateObject("ADODB.Stream")
		DoteyUpload_SourceData.Type = 1 ' 二进制流
		DoteyUpload_SourceData.Open

		Version = "1.0 Beta"	' 版本
		ErrMsg = ""	' 错误信息
		Set Progress = New ProgressInfo

	End Sub

	' 将文件根据其文件名统一保存在某路径下
	Public Sub SaveTo(path)
		
		Upload()	' 上传

		if right(path,1) <> "/" then path = path & "/" 

		' 遍历所有已上传文件
		For Each fileItem In Files.Items			
			fileItem.SaveAs path & fileItem.FileName
		Next

		' 保存结束后更新进度信息
		Progress.ReadyState = "complete" '上传结束
		UpdateProgressInfo progressID

	End Sub

	' 分析上传的数据，并保存到相应集合中
	Public Sub Upload ()

		Dim TotalBytes, Boundary
		TotalBytes = Request.TotalBytes	 ' 总大小
		If TotalBytes < 1 Then
			Raise("无数据传入")
			Exit Sub
		End If
		If TotalBytes > MaxTotalBytes Then
			Raise("您当前上传大小为" & TotalBytes/1000 & " K，最大允许为" & MaxTotalBytes/1024 & "K")
			Exit Sub
		End If
		Boundary = GetBoundary()
		If IsNull(Boundary) Then 
			Raise("如果form中没有包括multipart/form-data上传是无效的")
			Exit Sub	 ''如果form中没有包括multipart/form-data上传是无效的
		End If
		Boundary = StringToBinary(Boundary)
		
		Progress.ReadyState = "loading" '开始上传
		Progress.TotalBytes = TotalBytes
		UpdateProgressInfo progressID

		Dim DataPart, PartSize
		BytesRead = 0

		'循环分块读取
		Do While BytesRead < TotalBytes

			'分块读取
			PartSize = ChunkReadSize
			if PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize

			DoteyUpload_SourceData.Write DataPart

			Progress.UploadedBytes = BytesRead
			Progress.LastActivity = Now()

			' 更新进度信息
			UpdateProgressInfo progressID

		Loop

		' 上传结束后更新进度信息
		Progress.ReadyState = "loaded" '上传结束
		UpdateProgressInfo progressID

		Dim Binary
		DoteyUpload_SourceData.Position = 0
		Binary = DoteyUpload_SourceData.Read

		Dim BoundaryStart, BoundaryEnd, PosEndOfHeader, IsBoundaryEnd
		Dim Header, bFieldContent
		Dim FieldName
		Dim File
		Dim TwoCharsAfterEndBoundary

		BoundaryStart = InStrB(Binary, Boundary)
		BoundaryEnd = InStrB(BoundaryStart + LenB(Boundary), Binary, Boundary, 0)

		Do While (BoundaryStart > 0 And BoundaryEnd > 0 And Not IsBoundaryEnd)
			' 获取表单头的结束位置
			PosEndOfHeader = InStrB(BoundaryStart + LenB(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))
						
			' 分离表单头信息，类似于：
			' Content-Disposition: form-data; name="file1"; filename="G:\homepage.txt"
			' Content-Type: text/plain
			Header = BinaryToString(MidB(Binary, BoundaryStart + LenB(Boundary) + 2, PosEndOfHeader - BoundaryStart - LenB(Boundary) - 2))

			' 分离表单内容
			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), BoundaryEnd - (PosEndOfHeader + 4) - 2)
			
			FieldName = GetFieldName(Header)
			' 如果是附件
			If InStr (Header,"filename=""") > 0 Then
				Set File = New FileInfo
				
				' 获取文件相关信息
				Dim clientPath
				clientPath = GetFileName(Header)
				File.FileName = GetFileNameByPath(clientPath)
				File.FileExt = GetFileExt(clientPath)
				File.FilePath = clientPath
				File.FileType = GetFileType(Header)
				File.FileStart = PosEndOfHeader + 3
				File.FileSize = BoundaryEnd - (PosEndOfHeader + 4) - 2
				File.FormName = FieldName

				' 如果该文件不为空并不存在该表单项保存之
				If Not Files.Exists(FieldName) And File.FileSize > 0 Then
					Files.Add FieldName, File
				End If
			'表单数据				
			Else
				' 允许同名表单
				If Form.Exists(FieldName) Then
					Form(FieldName) = Form(FieldName) & "," & BinaryToString(bFieldContent)
				Else
					Form.Add FieldName, BinaryToString(bFieldContent)
				End If
			End If

			' 是否结束位置
			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, BoundaryEnd + LenB(Boundary), 2))
			IsBoundaryEnd = TwoCharsAfterEndBoundary = "--"

			If Not IsBoundaryEnd Then ' 如果不是结尾, 继续读取下一块
				BoundaryStart = BoundaryEnd
				BoundaryEnd = InStrB(BoundaryStart + LenB(Boundary), Binary, Boundary)
			End If
		Loop
		
		' 解析文件结束后更新进度信息
		Progress.UploadedBytes = TotalBytes
		Progress.ReadyState = "interactive" '解析文件结束
		UpdateProgressInfo progressID

	End Sub

	'异常信息
	Private Sub Raise(Message)
		ErrMsg = ErrMsg & "[" & Now & "]" & Message & "<BR>"
		
		Progress.ErrorMessage = Message
		UpdateProgressInfo ProgressID
		
		'call Err.Raise(vbObjectError, "DoteyUpload", Message)

	End Sub

	' 取边界值
	Private Function GetBoundary()
		Dim ContentType, ctArray, bArray
		ContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
		ctArray = Split(ContentType, ";")
		If Trim(ctArray(0)) = "multipart/form-data" Then
			bArray = Split(Trim(ctArray(1)), "=")
			GetBoundary = "--" & Trim(bArray(1))
		Else	'如果form中没有包括multipart/form-data上传是无效的
			GetBoundary = null
			Raise("如果form中没有包括multipart/form-data上传是无效的")
		End If
	End Function

	' 将二进制流转化成文本
	Private Function BinaryToString(xBinary)
		Dim Binary
		if vartype(xBinary) = 8 then Binary = MultiByteToBinary(xBinary) else Binary = xBinary
		
	  Dim RS, LBinary
	  Const adLongVarChar = 201
	  Set RS = CreateObject("ADODB.Recordset")
	  LBinary = LenB(Binary)
		
		if LBinary>0 then
			RS.Fields.Append "mBinary", adLongVarChar, LBinary
			RS.Open
			RS.AddNew
				RS("mBinary").AppendChunk Binary 
			RS.Update
			BinaryToString = RS("mBinary")
		Else
			BinaryToString = ""
		End If
	End Function


	Function MultiByteToBinary(MultiByte)
	  Dim RS, LMultiByte, Binary
	  Const adLongVarBinary = 205
	  Set RS = CreateObject("ADODB.Recordset")
	  LMultiByte = LenB(MultiByte)
		if LMultiByte>0 then
			RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
			RS.Open
			RS.AddNew
				RS("mBinary").AppendChunk MultiByte & ChrB(0)
			RS.Update
			Binary = RS("mBinary").GetChunk(LMultiByte)
		End If
	  MultiByteToBinary = Binary
	End Function


	' 字符串到二进制
	Function StringToBinary(String)
		Dim I, B
		For I=1 to len(String)
			B = B & ChrB(Asc(Mid(String,I,1)))
		Next
		StringToBinary = B
	End Function

	'返回表单名
	Private Function GetFieldName(infoStr)
		Dim sPos, EndPos
		sPos = InStr(infoStr, "name=")
		EndPos = InStr(sPos + 6, infoStr, Chr(34) & ";")
		If EndPos = 0 Then
			EndPos = inStr(sPos + 6, infoStr, Chr(34))
		End If
		GetFieldName = Mid(infoStr, sPos + 6, endPos - _
			(sPos + 6))
	End Function

	'返回文件名
	Private Function GetFileName(infoStr)
		Dim sPos, EndPos
		sPos = InStr(infoStr, "filename=")
		EndPos = InStr(infoStr, Chr(34) & CrLf)
		GetFileName = Mid(infoStr, sPos + 10, EndPos - _
			(sPos + 10))
	End Function

	'返回文件的 MIME type
	Private Function GetFileType(infoStr)
		sPos = InStr(infoStr, "Content-Type: ")
		GetFileType = Mid(infoStr, sPos + 14)
	End Function

	'根据路径获取文件名
	Private Function GetFileNameByPath(FullPath)
		Dim pos
		pos = 0
		FullPath = Replace(FullPath, "/", "\")
		pos = InStrRev(FullPath, "\") + 1
		If (pos > 0) Then
			GetFileNameByPath = Mid(FullPath, pos)
		Else
			GetFileNameByPath = FullPath
		End If
	End Function

	'根据路径获取扩展名
	Private Function GetFileExt(FullPath)
		Dim pos
		pos = InStrRev(FullPath,".")
		if pos>0 then GetFileExt = Mid(FullPath, Pos)
	End Function

	' 更新进度信息
	' 进度信息保存在Application中的ADODB.Recordset对象中
	Private Sub UpdateProgressInfo(progressID)
		Const adTypeText = 2, adDate = 7, adUnsignedInt = 19, adVarChar = 200
		
		If (progressID <> "" And IsNumeric(progressID)) Then
			Application.Lock()
			if IsEmpty(Application(UploadProgressInfo)) Then
				Set Info = Server.CreateObject("ADODB.Recordset")
				Set Application(UploadProgressInfo) = Info
				Info.Fields.Append "ProgressID", adUnsignedInt
				Info.Fields.Append "StartTime", adDate
				Info.Fields.Append "LastActivity", adDate
				Info.Fields.Append "TotalBytes", adUnsignedInt
				Info.Fields.Append "UploadedBytes", adUnsignedInt
				Info.Fields.Append "ReadyState", adVarChar, 128
				Info.Fields.Append "ErrorMessage", adVarChar, 4000
				Info.Open 
		 		Info("ProgressID").Properties("Optimize") = true
				Info.AddNew 
			Else
				Set Info = Application(UploadProgressInfo)
				If Not Info.Eof Then
					Info.MoveFirst()
					Info.Find "ProgressID = " & progressID
				End If
				If (Info.EOF) Then
					Info.AddNew
				End If
			End If

			Info("ProgressID") = clng(progressID)
			Info("StartTime") = Progress.StartTime
			Info("LastActivity") = Now()
			Info("TotalBytes") = Progress.TotalBytes
			Info("UploadedBytes") = Progress.UploadedBytes
			Info("ReadyState") = Progress.ReadyState
			Info("ErrorMessage") = Progress.ErrorMessage
			Info.Update

			Application.UnLock
		End IF
	End Sub

	' 根据上传ID获取进度信息
	Public Function GetProgressInfo(progressID)

		Dim pi, Infos
		Set pi = New ProgressInfo
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Set Infos = Application(UploadProgressInfo)
			If Not Infos.Eof Then
				Infos.MoveFirst
				Infos.Find "ProgressID = " & progressID
				If Not Infos.EOF Then
					pi.StartTime = Infos("StartTime")
					pi.LastActivity = Infos("LastActivity")
					pi.TotalBytes = clng(Infos("TotalBytes"))
					pi.UploadedBytes = clng(Infos("UploadedBytes"))
					pi.ReadyState = Trim(Infos("ReadyState"))
					pi.ErrorMessage = Trim(Infos("ErrorMessage"))
					Set GetProgressInfo = pi
				End If
			End If
		End If
		Set GetProgressInfo = pi
	End Function

	' 移除指定的进度信息
	Private Sub RemoveProgressInfo(progressID)
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Application.Lock
			Set Info = Application(UploadProgressInfo)
			If Not Info.Eof Then
				Info.MoveFirst
				Info.Find "ProgressID = " & progressID
				If  Not Info.EOF Then
					Info.Delete
				End If
			End If

			' 如果没有记录了, 直接释放, 避免'800a0bcd'错误
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

	' 移除指定的进度信息
	Private Sub RemoveOldProgressInfo(progressID)
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Dim L
			Application.Lock

			Set Info = Application(UploadProgressInfo)
			Info.MoveFirst

			Do
				L = Info("LastActivity").Value
				If IsEmpty(L) Then
					Info.Delete() 
				ElseIf DateDiff("d", Now(), L) > 30 Then
					Info.Delete()
				End If
				Info.MoveNext()
			Loop Until Info.EOF

			' 如果没有记录了, 直接释放, 避免'800a0bcd'错误
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

End Class

'---------------------------------------------------
' 进度信息 类
'---------------------------------------------------
Class ProgressInfo
	
	Public UploadedBytes
	Public TotalBytes
	Public StartTime
	Public LastActivity
	Public ReadyState
	Public ErrorMessage

	Private Sub Class_Initialize()
		UploadedBytes = 0	' 已上传大小
		TotalBytes = 0	' 总大小
		StartTime = Now()	' 开始时间
		LastActivity = Now()	 ' 最后更新时间
		ReadyState = "uninitialized"	' uninitialized,loading,loaded,interactive,complete
		ErrorMessage = ""
	End Sub

	' 总大小
	Public Property Get TotalSize
		TotalSize = FormatNumber(TotalBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' 已上传大小
	Public Property Get SizeCompleted
		SizeCompleted = FormatNumber(UploadedBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' 已上传秒数
	Public Property Get ElapsedSeconds
		ElapsedSeconds = DateDiff("s", StartTime, Now())
	End Property 

	' 已上传时间
	Public Property Get ElapsedTime
		If ElapsedSeconds > 3600 then
			ElapsedTime = ElapsedSeconds \ 3600 & " 时 " & (ElapsedSeconds mod 3600) \ 60 & " 分 " & ElapsedSeconds mod 60 & " 秒"
		ElseIf ElapsedSeconds > 60 then
			ElapsedTime = ElapsedSeconds \ 60 & " 分 " & ElapsedSeconds mod 60 & " 秒"
		else
			ElapsedTime = ElapsedSeconds mod 60 & " 秒"
		End If
	End Property 

	' 传输速率
	Public Property Get TransferRate
		If ElapsedSeconds > 0 Then
			TransferRate = FormatNumber(UploadedBytes / 1024 / ElapsedSeconds, 2, 0, 0, -1) & " K/秒"
		Else
			TransferRate = "0 K/秒"
		End If
	End Property 

	' 完成百分比
	Public Property Get Percentage
		If TotalBytes > 0 Then
			Percentage = fix(UploadedBytes / TotalBytes * 100) & "%"
		Else
			Percentage = "0%"
		End If
	End Property 

	' 估计剩余时间
	Public Property Get TimeLeft
		If UploadedBytes > 0 Then
			SecondsLeft = fix(ElapsedSeconds * (TotalBytes / UploadedBytes - 1))
			If SecondsLeft > 3600 then
				TimeLeft = SecondsLeft \ 3600 & " 时 " & (SecondsLeft mod 3600) \ 60 & " 分 " & SecondsLeft mod 60 & " 秒"
			ElseIf SecondsLeft > 60 then
				TimeLeft = SecondsLeft \ 60 & " 分 " & SecondsLeft mod 60 & " 秒"
			else
				TimeLeft = SecondsLeft mod 60 & " 秒"
			End If
		Else
			TimeLeft = "未知"
		End If
	End Property 

End Class

'---------------------------------------------------
' 文件信息 类
'---------------------------------------------------
Class FileInfo
	
	Dim FormName, FileName, FilePath, FileSize, FileType, FileStart, FileExt, NewFileName

	Private Sub Class_Initialize 
		FileName = ""		' 文件名
		FilePath = ""			' 客户端路径
		FileSize = 0			' 文件大小
		FileStart= 0			' 文件开始位置
		FormName = ""	' 表单名
		FileType = ""		' 文件Content Type
		FileExt = ""			' 文件扩展名
		NewFileName = ""	'上传后文件名
	End Sub

	Public Function Save()
		SaveAs(FileName)
	End Function

	' 保存文件
	Public Function SaveAs(fullpath)
		Dim dr
		SaveAs = false
		If trim(fullpath) = "" Or FileStart = 0 Or FileName = "" Or right(fullpath,1) = "/" Then Exit Function
		
		NewFileName = GetFileNameByPath(fullpath)

		Set dr = CreateObject("Adodb.Stream")
		dr.Mode = 3
		dr.Type = 1
		dr.Open
		DoteyUpload_SourceData.position = FileStart
		DoteyUpload_SourceData.copyto dr, FileSize
		dr.SaveToFile MapPath(FullPath), 2
		dr.Close
		set dr = nothing 
		SaveAs = true
	End function

	' 返回Binary
	Public Function GetBinary()
		Dim Binary
		If FileStart = 0 Then Exit Function

		DoteyUpload_SourceData.Position = FileStart
		Binary = DoteyUpload_SourceData.Read(FileSize)

		GetBinary = Binary
	End function

	' 取服务器端路径
	Private Function MapPath(Path)
		If InStr(1, Path, ":") > 0 Or Left(Path, 2) = "\\" Then
			MapPath = Path 
		Else 
			MapPath = Server.MapPath(Path)
		End If
	End function

	'根据路径获取文件名
	Private Function GetFileNameByPath(FullPath)
		Dim pos
		pos = 0
		FullPath = Replace(FullPath, "/", "\")
		pos = InStrRev(FullPath, "\") + 1
		If (pos > 0) Then
			GetFileNameByPath = Mid(FullPath, pos)
		Else
			GetFileNameByPath = FullPath
		End If
	End Function

End Class
'============================================================
Class DictionaryClass
Dim ArryObj()     '使用该二维数组来做存放数据的字典
Dim MaxIndex       'MaxIndex则是ArryObj开始的最大上标
Dim CurIndex       '字典指针,用来指向ArryObj的指针
Dim C_ErrCode       '错误代码号

Private Sub Class_Initialize 
CurIndex=0       '从下标0开始
C_ErrCode=0       '0表示没有任何错误
MaxIndex=50       '默认的大小
Redim ArryObj(1,MaxIndex)   '定义一个二维的数组
End Sub

Private Sub Class_Terminate 
Erase ArryObj   '清除数组
End Sub

Public Property Get ErrCode '返回错误代码
  ErrCode=C_ErrCode
End Property

Public Property Get Count   '返回数据的总数,只返回CurIndex当前值-1即可.
  Count=CurIndex
End Property

Public Property Get Keys   '返回字典数据的全部Keys,返回数组.
  Dim KeyCount,ArryKey(),I
  KeyCount=CurIndex-1
  Redim ArryKey(KeyCount)
  For I=0 To KeyCount
    ArryKey(I)=ArryObj(0,I)
  Next
  Keys=ArryKey
  Erase ArryKey
End Property

Public Property Get Items   '返回字典数据的全部Values,返回数组.
  Dim KeyCount,ArryItem(),I
  KeyCount=CurIndex-1
  Redim ArryItem(KeyCount)
  For I=0 To KeyCount
    If isObject(ArryObj(1,I)) Then
      Set ArryItem(I)=ArryObj(1,I)
  Else
    ArryItem(I)=ArryObj(1,I)
  End If
  Next
  Items=ArryItem
  Erase ArryItem
End Property


Public Property Let Item(sKey,sVal) '取得sKey为Key的字典数据
  If sIsEmpty(sKey) Then
  Exit Property
  End If
  Dim i,iType
  iType=GetType(sKey)
  If iType=1 Then '如果sKey为数值型的则检查范围
  If sKey>CurIndex Or sKey<1 Then
  C_ErrCode=2
    Exit Property
End If
  End If
  If iType=0 Then
  For i=0 to CurIndex-1
    If ArryObj(0,i)=sKey Then
    If isObject(sVal) Then
      Set ArryObj(1,i)=sVal
  Else
    ArryObj(1,i)=sVal
  End If
  Exit Property
  End If
  Next
  ElseIf iType=1 Then
      sKey=sKey-1
    If isObject(sVal) Then
      Set ArryObj(1,sKey)=sVal
  Else
    ArryObj(1,sKey)=sVal
  End If
  Exit Property
  End If
  C_ErrCode=2         'ErrCode为2则是替换或个为sKey的字典数据时找不到数据
End Property

Public Property Get Item(sKey)
  If sIsEmpty(sKey) Then
    Item=Null
  Exit Property
  End If
  Dim i,iType
  iType=GetType(sKey)
  If iType=1 Then '如果sKey为数值型的则检查范围
  If sKey>CurIndex Or sKey<1 Then
    Item=Null
  Exit Property
End If
  End If
  If iType=0 Then
  For i=0 to CurIndex-1
    If ArryObj(0,i)=sKey Then
    If isObject(ArryObj(1,i)) Then
      Set Item=ArryObj(1,i)
  Else
    Item=ArryObj(1,i)
  End If
  Exit Property
  End If
  Next
  ElseIf iType=1 Then
      sKey=sKey-1
    If isObject(ArryObj(1,sKey)) Then
      Set Item=ArryObj(1,sKey)
  Else
    Item=ArryObj(1,sKey)
  End If
  Exit Property
  End If
  Item=Null
End Property

Public Sub Add(sKey,sVal) '添加字典
  'On Error Resume Next
  If Exists(sKey) Or C_ErrCode=9 Then
  C_ErrCode=1           'Key值不唯一(空的Key值也不能添加数字)
  Exit Sub
End If
  If CurIndex>MaxIndex Then
  MaxIndex=MaxIndex+1       '每次增加一个标数,可以按场合需求改为所需量
  Redim Preserve ArryObj(1,MaxIndex)
End If
ArryObj(0,CurIndex)=Cstr(sKey)     'sKey是标识值,将Key以字符串类型保存
if isObject(sVal) Then
  Set ArryObj(1,CurIndex)=sVal     'sVal是数据
Else
  ArryObj(1,CurIndex)=sVal     'sVal是数据
End If
CurIndex=CurIndex+1
End Sub

Public Sub Insert(sKey,nKey,nVal,sMethod)
If Not Exists(sKey) Then
  C_ErrCode=4
  Exit Sub
End If
  If Exists(nKey) Or C_ErrCode=9 Then
  C_ErrCode=4           'Key值不唯一(空的Key值也不能添加数字)
  Exit Sub
End If
sType=GetType(sKey)         '取得sKey的变量类型
Dim ArryResult(),I,sType,subIndex,sAdd
ReDim ArryResult(1,CurIndex)   '定义一个数组用来做临时存放地
if sIsEmpty(sMethod) Then sMethod="b"   '为空的数据则默认是"b"
sMethod=lcase(cstr(sMethod))
subIndex=CurIndex-1
sAdd=0
If sType=0 Then             '字符串类型比较
  If sMethod="1" Or sMethod="b" Or sMethod="back" Then '将数据插入sKey的后面
    For I=0 TO subIndex
      ArryResult(0,sAdd)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,sAdd)=ArryObj(1,I)
  Else
    ArryResult(1,sAdd)=ArryObj(1,I)
  End If
  If ArryObj(0,I)=sKey Then '插入数据
    sAdd=sAdd+1
    ArryResult(0,sAdd)=nKey
  If IsObject(nVal) Then
    Set ArryResult(1,sAdd)=nVal
  Else
    ArryResult(1,sAdd)=nVal
  End If
  End If
  sAdd=sAdd+1
  Next
  Else
    For I=0 TO subIndex
  If ArryObj(0,I)=sKey Then '插入数据
    ArryResult(0,sAdd)=nKey
  If IsObject(nVal) Then
    Set ArryResult(1,sAdd)=nVal
  Else
    ArryResult(1,sAdd)=nVal
  End If
  sAdd=sAdd+1
  End If
  ArryResult(0,sAdd)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,sAdd)=ArryObj(1,I)
  Else
    ArryResult(1,sAdd)=ArryObj(1,I)
  End If
  sAdd=sAdd+1
  Next
  End If
ElseIf sType=1 Then
  sKey=sKey-1             '减1是为了符合日常习惯(从1开始)
  If sMethod="1" Or sMethod="b" Or sMethod="back" Then '将数据插入sKey的后面
    For I=0 TO sKey         '取sKey前面部分数据
      ArryResult(0,I)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,I)=ArryObj(1,I)
  Else
    ArryResult(1,I)=ArryObj(1,I)
  End If
  Next
'插入新的数据
ArryResult(0,sKey+1)=nKey
If IsObject(nVal) Then
  Set ArryResult(1,sKey+1)=nVal
Else
  ArryResult(1,sKey+1)=nVal
End If
'取sKey后面的数据
    For I=sKey+1 TO subIndex
      ArryResult(0,I+1)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,I+1)=ArryObj(1,I)
  Else
    ArryResult(1,I+1)=ArryObj(1,I)
  End If
  Next
  Else
    For I=0 TO sKey-1         '取sKey-1前面部分数据
      ArryResult(0,I)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,I)=ArryObj(1,I)
  Else
    ArryResult(1,I)=ArryObj(1,I)
  End If
  Next
'插入新的数据
ArryResult(0,sKey)=nKey
If IsObject(nVal) Then
  Set ArryResult(1,sKey)=nVal
Else
  ArryResult(1,sKey)=nVal
End If
'取sKey后面的数据
    For I=sKey TO subIndex
      ArryResult(0,I+1)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,I+1)=ArryObj(1,I)
  Else
    ArryResult(1,I+1)=ArryObj(1,I)
  End If
  Next
  End If
Else
  C_ErrCode=3
  Exit Sub
End If
ReDim ArryObj(1,CurIndex) '重置数据
For I=0 To CurIndex
ArryObj(0,I)=ArryResult(0,I)
If isObject(ArryResult(1,I)) Then
  Set ArryObj(1,I)=ArryResult(1,I)
Else
  ArryObj(1,I)=ArryResult(1,I)
End If
Next
MaxIndex=CurIndex
Erase ArryResult
CurIndex=CurIndex+1     'Insert后数据指针加一
End Sub

Public Function Exists(sKey)   '判断存不存在某个字典数据
  If sIsEmpty(sKey) Then
    Exists=False
    Exit Function
End If
Dim I,vType
vType=GetType(sKey)
If vType=0 Then
  For I=0 To CurIndex-1
  If ArryObj(0,I)=sKey Then
  Exists=True
  Exit Function
End If
  Next
ElseIf vType=1 Then
    If sKey<=CurIndex And sKey>0 Then
    Exists=True
    Exit Function
  End If
End If
Exists=False
End Function

Public Sub Remove(sKey)         '根据sKey的值Remove一条字典数据
If Not Exists(sKey) Then
  C_ErrCode=3
  Exit Sub
End If
sType=GetType(sKey)         '取得sKey的变量类型
Dim ArryResult(),I,sType,sAdd
ReDim ArryResult(1,CurIndex-2)   '定义一个数组用来做临时存放地
sAdd=0
If sType=0 Then             '字符串类型比较
    For I=0 TO CurIndex-1
  If ArryObj(0,I)<>sKey Then
      ArryResult(0,sAdd)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,sAdd)=ArryObj(1,I)
  Else
    ArryResult(1,sAdd)=ArryObj(1,I)
  End If
  sAdd=sAdd+1
End If
  Next
ElseIf sType=1 Then
  sKey=sKey-1             '减1是为了符合日常习惯(从1开始)
    For I=0 TO CurIndex-1
  If I<>sKey Then
      ArryResult(0,sAdd)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,sAdd)=ArryObj(1,I)
  Else
    ArryResult(1,sAdd)=ArryObj(1,I)
  End If
  sAdd=sAdd+1
End If
  Next
Else
  C_ErrCode=3
  Exit Sub
End If
MaxIndex=CurIndex-2
ReDim ArryObj(1,MaxIndex) '重置数据
For I=0 To MaxIndex
ArryObj(0,I)=ArryResult(0,I)
If isObject(ArryResult(1,I)) Then
  Set ArryObj(1,I)=ArryResult(1,I)
Else
  ArryObj(1,I)=ArryResult(1,I)
End If
Next
Erase ArryResult
CurIndex=CurIndex-1     '减一是Remove后数据指针
End Sub

Public Sub RemoveAll '全部清空字典数据,只Redim一下就OK了
  Redim ArryObj(MaxIndex)
CurIndex=0
End Sub

Public Sub ClearErr   '重置错误
  C_ErrCode=0
End Sub

Private Function sIsEmpty(sVal) '判断sVal是否为空值
  If IsEmpty(sVal) Then
  C_ErrCode=9           'Key值为空的错误代码
  sIsEmpty=True
  Exit Function
End If
  If IsNull(sVal) Then
  C_ErrCode=9           'Key值为空的错误代码
  sIsEmpty=True
  Exit Function
End If
  If Trim(sVal)="" Then
  C_ErrCode=9           'Key值为空的错误代码
  sIsEmpty=True
  Exit Function
End If
sIsEmpty=False
End Function

Private Function GetType(sVal)   '取得变量sVal的变量类型
  dim sType
sType=TypeName(sVal)
  Select Case sType
  Case "String"
    GetType=0
  Case "Integer","Long","Single","Double"
    GetType=1
  Case Else
    GetType=-1
End Select
End Function

End Class
%>
