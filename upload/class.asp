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
		Set Files = Server.CreateObject("Scripting.Dictionary")	' �ϴ��ļ�����
		Set Form = Server.CreateObject("Scripting.Dictionary")	' ������
'		Set Files = new DictionaryClass	' �ϴ��ļ�����
'		Set Form = new DictionaryClass	' ������
		UploadProgressInfo = "DoteyUploadProgressInfo"  ' Application��Key
		MaxTotalBytes = 1 *1024 *1024 *1024 ' Ĭ�����1G
		ChunkReadSize = 64 * 1024	' �ֿ��С64K
		CrLf = Chr(13) & Chr(10)	' ����

		Set DoteyUpload_SourceData = Server.CreateObject("ADODB.Stream")
		DoteyUpload_SourceData.Type = 1 ' ��������
		DoteyUpload_SourceData.Open

		Version = "1.0 Beta"	' �汾
		ErrMsg = ""	' ������Ϣ
		Set Progress = New ProgressInfo

	End Sub

	' ���ļ��������ļ���ͳһ������ĳ·����
	Public Sub SaveTo(path)
		
		Upload()	' �ϴ�

		if right(path,1) <> "/" then path = path & "/" 

		' �����������ϴ��ļ�
		For Each fileItem In Files.Items			
			fileItem.SaveAs path & fileItem.FileName
		Next

		' �����������½�����Ϣ
		Progress.ReadyState = "complete" '�ϴ�����
		UpdateProgressInfo progressID

	End Sub

	' �����ϴ������ݣ������浽��Ӧ������
	Public Sub Upload ()

		Dim TotalBytes, Boundary
		TotalBytes = Request.TotalBytes	 ' �ܴ�С
		If TotalBytes < 1 Then
			Raise("�����ݴ���")
			Exit Sub
		End If
		If TotalBytes > MaxTotalBytes Then
			Raise("����ǰ�ϴ���СΪ" & TotalBytes/1000 & " K���������Ϊ" & MaxTotalBytes/1024 & "K")
			Exit Sub
		End If
		Boundary = GetBoundary()
		If IsNull(Boundary) Then 
			Raise("���form��û�а���multipart/form-data�ϴ�����Ч��")
			Exit Sub	 ''���form��û�а���multipart/form-data�ϴ�����Ч��
		End If
		Boundary = StringToBinary(Boundary)
		
		Progress.ReadyState = "loading" '��ʼ�ϴ�
		Progress.TotalBytes = TotalBytes
		UpdateProgressInfo progressID

		Dim DataPart, PartSize
		BytesRead = 0

		'ѭ���ֿ��ȡ
		Do While BytesRead < TotalBytes

			'�ֿ��ȡ
			PartSize = ChunkReadSize
			if PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize

			DoteyUpload_SourceData.Write DataPart

			Progress.UploadedBytes = BytesRead
			Progress.LastActivity = Now()

			' ���½�����Ϣ
			UpdateProgressInfo progressID

		Loop

		' �ϴ���������½�����Ϣ
		Progress.ReadyState = "loaded" '�ϴ�����
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
			' ��ȡ��ͷ�Ľ���λ��
			PosEndOfHeader = InStrB(BoundaryStart + LenB(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))
						
			' �����ͷ��Ϣ�������ڣ�
			' Content-Disposition: form-data; name="file1"; filename="G:\homepage.txt"
			' Content-Type: text/plain
			Header = BinaryToString(MidB(Binary, BoundaryStart + LenB(Boundary) + 2, PosEndOfHeader - BoundaryStart - LenB(Boundary) - 2))

			' ���������
			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), BoundaryEnd - (PosEndOfHeader + 4) - 2)
			
			FieldName = GetFieldName(Header)
			' ����Ǹ���
			If InStr (Header,"filename=""") > 0 Then
				Set File = New FileInfo
				
				' ��ȡ�ļ������Ϣ
				Dim clientPath
				clientPath = GetFileName(Header)
				File.FileName = GetFileNameByPath(clientPath)
				File.FileExt = GetFileExt(clientPath)
				File.FilePath = clientPath
				File.FileType = GetFileType(Header)
				File.FileStart = PosEndOfHeader + 3
				File.FileSize = BoundaryEnd - (PosEndOfHeader + 4) - 2
				File.FormName = FieldName

				' ������ļ���Ϊ�ղ������ڸñ����֮
				If Not Files.Exists(FieldName) And File.FileSize > 0 Then
					Files.Add FieldName, File
				End If
			'������				
			Else
				' ����ͬ����
				If Form.Exists(FieldName) Then
					Form(FieldName) = Form(FieldName) & "," & BinaryToString(bFieldContent)
				Else
					Form.Add FieldName, BinaryToString(bFieldContent)
				End If
			End If

			' �Ƿ����λ��
			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, BoundaryEnd + LenB(Boundary), 2))
			IsBoundaryEnd = TwoCharsAfterEndBoundary = "--"

			If Not IsBoundaryEnd Then ' ������ǽ�β, ������ȡ��һ��
				BoundaryStart = BoundaryEnd
				BoundaryEnd = InStrB(BoundaryStart + LenB(Boundary), Binary, Boundary)
			End If
		Loop
		
		' �����ļ���������½�����Ϣ
		Progress.UploadedBytes = TotalBytes
		Progress.ReadyState = "interactive" '�����ļ�����
		UpdateProgressInfo progressID

	End Sub

	'�쳣��Ϣ
	Private Sub Raise(Message)
		ErrMsg = ErrMsg & "[" & Now & "]" & Message & "<BR>"
		
		Progress.ErrorMessage = Message
		UpdateProgressInfo ProgressID
		
		'call Err.Raise(vbObjectError, "DoteyUpload", Message)

	End Sub

	' ȡ�߽�ֵ
	Private Function GetBoundary()
		Dim ContentType, ctArray, bArray
		ContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
		ctArray = Split(ContentType, ";")
		If Trim(ctArray(0)) = "multipart/form-data" Then
			bArray = Split(Trim(ctArray(1)), "=")
			GetBoundary = "--" & Trim(bArray(1))
		Else	'���form��û�а���multipart/form-data�ϴ�����Ч��
			GetBoundary = null
			Raise("���form��û�а���multipart/form-data�ϴ�����Ч��")
		End If
	End Function

	' ����������ת�����ı�
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


	' �ַ�����������
	Function StringToBinary(String)
		Dim I, B
		For I=1 to len(String)
			B = B & ChrB(Asc(Mid(String,I,1)))
		Next
		StringToBinary = B
	End Function

	'���ر���
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

	'�����ļ���
	Private Function GetFileName(infoStr)
		Dim sPos, EndPos
		sPos = InStr(infoStr, "filename=")
		EndPos = InStr(infoStr, Chr(34) & CrLf)
		GetFileName = Mid(infoStr, sPos + 10, EndPos - _
			(sPos + 10))
	End Function

	'�����ļ��� MIME type
	Private Function GetFileType(infoStr)
		sPos = InStr(infoStr, "Content-Type: ")
		GetFileType = Mid(infoStr, sPos + 14)
	End Function

	'����·����ȡ�ļ���
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

	'����·����ȡ��չ��
	Private Function GetFileExt(FullPath)
		Dim pos
		pos = InStrRev(FullPath,".")
		if pos>0 then GetFileExt = Mid(FullPath, Pos)
	End Function

	' ���½�����Ϣ
	' ������Ϣ������Application�е�ADODB.Recordset������
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

	' �����ϴ�ID��ȡ������Ϣ
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

	' �Ƴ�ָ���Ľ�����Ϣ
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

			' ���û�м�¼��, ֱ���ͷ�, ����'800a0bcd'����
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

	' �Ƴ�ָ���Ľ�����Ϣ
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

			' ���û�м�¼��, ֱ���ͷ�, ����'800a0bcd'����
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

End Class

'---------------------------------------------------
' ������Ϣ ��
'---------------------------------------------------
Class ProgressInfo
	
	Public UploadedBytes
	Public TotalBytes
	Public StartTime
	Public LastActivity
	Public ReadyState
	Public ErrorMessage

	Private Sub Class_Initialize()
		UploadedBytes = 0	' ���ϴ���С
		TotalBytes = 0	' �ܴ�С
		StartTime = Now()	' ��ʼʱ��
		LastActivity = Now()	 ' ������ʱ��
		ReadyState = "uninitialized"	' uninitialized,loading,loaded,interactive,complete
		ErrorMessage = ""
	End Sub

	' �ܴ�С
	Public Property Get TotalSize
		TotalSize = FormatNumber(TotalBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' ���ϴ���С
	Public Property Get SizeCompleted
		SizeCompleted = FormatNumber(UploadedBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' ���ϴ�����
	Public Property Get ElapsedSeconds
		ElapsedSeconds = DateDiff("s", StartTime, Now())
	End Property 

	' ���ϴ�ʱ��
	Public Property Get ElapsedTime
		If ElapsedSeconds > 3600 then
			ElapsedTime = ElapsedSeconds \ 3600 & " ʱ " & (ElapsedSeconds mod 3600) \ 60 & " �� " & ElapsedSeconds mod 60 & " ��"
		ElseIf ElapsedSeconds > 60 then
			ElapsedTime = ElapsedSeconds \ 60 & " �� " & ElapsedSeconds mod 60 & " ��"
		else
			ElapsedTime = ElapsedSeconds mod 60 & " ��"
		End If
	End Property 

	' ��������
	Public Property Get TransferRate
		If ElapsedSeconds > 0 Then
			TransferRate = FormatNumber(UploadedBytes / 1024 / ElapsedSeconds, 2, 0, 0, -1) & " K/��"
		Else
			TransferRate = "0 K/��"
		End If
	End Property 

	' ��ɰٷֱ�
	Public Property Get Percentage
		If TotalBytes > 0 Then
			Percentage = fix(UploadedBytes / TotalBytes * 100) & "%"
		Else
			Percentage = "0%"
		End If
	End Property 

	' ����ʣ��ʱ��
	Public Property Get TimeLeft
		If UploadedBytes > 0 Then
			SecondsLeft = fix(ElapsedSeconds * (TotalBytes / UploadedBytes - 1))
			If SecondsLeft > 3600 then
				TimeLeft = SecondsLeft \ 3600 & " ʱ " & (SecondsLeft mod 3600) \ 60 & " �� " & SecondsLeft mod 60 & " ��"
			ElseIf SecondsLeft > 60 then
				TimeLeft = SecondsLeft \ 60 & " �� " & SecondsLeft mod 60 & " ��"
			else
				TimeLeft = SecondsLeft mod 60 & " ��"
			End If
		Else
			TimeLeft = "δ֪"
		End If
	End Property 

End Class

'---------------------------------------------------
' �ļ���Ϣ ��
'---------------------------------------------------
Class FileInfo
	
	Dim FormName, FileName, FilePath, FileSize, FileType, FileStart, FileExt, NewFileName

	Private Sub Class_Initialize 
		FileName = ""		' �ļ���
		FilePath = ""			' �ͻ���·��
		FileSize = 0			' �ļ���С
		FileStart= 0			' �ļ���ʼλ��
		FormName = ""	' ����
		FileType = ""		' �ļ�Content Type
		FileExt = ""			' �ļ���չ��
		NewFileName = ""	'�ϴ����ļ���
	End Sub

	Public Function Save()
		SaveAs(FileName)
	End Function

	' �����ļ�
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

	' ����Binary
	Public Function GetBinary()
		Dim Binary
		If FileStart = 0 Then Exit Function

		DoteyUpload_SourceData.Position = FileStart
		Binary = DoteyUpload_SourceData.Read(FileSize)

		GetBinary = Binary
	End function

	' ȡ��������·��
	Private Function MapPath(Path)
		If InStr(1, Path, ":") > 0 Or Left(Path, 2) = "\\" Then
			MapPath = Path 
		Else 
			MapPath = Server.MapPath(Path)
		End If
	End function

	'����·����ȡ�ļ���
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
Dim ArryObj()     'ʹ�øö�ά��������������ݵ��ֵ�
Dim MaxIndex       'MaxIndex����ArryObj��ʼ������ϱ�
Dim CurIndex       '�ֵ�ָ��,����ָ��ArryObj��ָ��
Dim C_ErrCode       '��������

Private Sub Class_Initialize 
CurIndex=0       '���±�0��ʼ
C_ErrCode=0       '0��ʾû���κδ���
MaxIndex=50       'Ĭ�ϵĴ�С
Redim ArryObj(1,MaxIndex)   '����һ����ά������
End Sub

Private Sub Class_Terminate 
Erase ArryObj   '�������
End Sub

Public Property Get ErrCode '���ش������
  ErrCode=C_ErrCode
End Property

Public Property Get Count   '�������ݵ�����,ֻ����CurIndex��ǰֵ-1����.
  Count=CurIndex
End Property

Public Property Get Keys   '�����ֵ����ݵ�ȫ��Keys,��������.
  Dim KeyCount,ArryKey(),I
  KeyCount=CurIndex-1
  Redim ArryKey(KeyCount)
  For I=0 To KeyCount
    ArryKey(I)=ArryObj(0,I)
  Next
  Keys=ArryKey
  Erase ArryKey
End Property

Public Property Get Items   '�����ֵ����ݵ�ȫ��Values,��������.
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


Public Property Let Item(sKey,sVal) 'ȡ��sKeyΪKey���ֵ�����
  If sIsEmpty(sKey) Then
  Exit Property
  End If
  Dim i,iType
  iType=GetType(sKey)
  If iType=1 Then '���sKeyΪ��ֵ�͵����鷶Χ
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
  C_ErrCode=2         'ErrCodeΪ2�����滻���ΪsKey���ֵ�����ʱ�Ҳ�������
End Property

Public Property Get Item(sKey)
  If sIsEmpty(sKey) Then
    Item=Null
  Exit Property
  End If
  Dim i,iType
  iType=GetType(sKey)
  If iType=1 Then '���sKeyΪ��ֵ�͵����鷶Χ
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

Public Sub Add(sKey,sVal) '����ֵ�
  'On Error Resume Next
  If Exists(sKey) Or C_ErrCode=9 Then
  C_ErrCode=1           'Keyֵ��Ψһ(�յ�KeyֵҲ�����������)
  Exit Sub
End If
  If CurIndex>MaxIndex Then
  MaxIndex=MaxIndex+1       'ÿ������һ������,���԰����������Ϊ������
  Redim Preserve ArryObj(1,MaxIndex)
End If
ArryObj(0,CurIndex)=Cstr(sKey)     'sKey�Ǳ�ʶֵ,��Key���ַ������ͱ���
if isObject(sVal) Then
  Set ArryObj(1,CurIndex)=sVal     'sVal������
Else
  ArryObj(1,CurIndex)=sVal     'sVal������
End If
CurIndex=CurIndex+1
End Sub

Public Sub Insert(sKey,nKey,nVal,sMethod)
If Not Exists(sKey) Then
  C_ErrCode=4
  Exit Sub
End If
  If Exists(nKey) Or C_ErrCode=9 Then
  C_ErrCode=4           'Keyֵ��Ψһ(�յ�KeyֵҲ�����������)
  Exit Sub
End If
sType=GetType(sKey)         'ȡ��sKey�ı�������
Dim ArryResult(),I,sType,subIndex,sAdd
ReDim ArryResult(1,CurIndex)   '����һ��������������ʱ��ŵ�
if sIsEmpty(sMethod) Then sMethod="b"   'Ϊ�յ�������Ĭ����"b"
sMethod=lcase(cstr(sMethod))
subIndex=CurIndex-1
sAdd=0
If sType=0 Then             '�ַ������ͱȽ�
  If sMethod="1" Or sMethod="b" Or sMethod="back" Then '�����ݲ���sKey�ĺ���
    For I=0 TO subIndex
      ArryResult(0,sAdd)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,sAdd)=ArryObj(1,I)
  Else
    ArryResult(1,sAdd)=ArryObj(1,I)
  End If
  If ArryObj(0,I)=sKey Then '��������
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
  If ArryObj(0,I)=sKey Then '��������
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
  sKey=sKey-1             '��1��Ϊ�˷����ճ�ϰ��(��1��ʼ)
  If sMethod="1" Or sMethod="b" Or sMethod="back" Then '�����ݲ���sKey�ĺ���
    For I=0 TO sKey         'ȡsKeyǰ�沿������
      ArryResult(0,I)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,I)=ArryObj(1,I)
  Else
    ArryResult(1,I)=ArryObj(1,I)
  End If
  Next
'�����µ�����
ArryResult(0,sKey+1)=nKey
If IsObject(nVal) Then
  Set ArryResult(1,sKey+1)=nVal
Else
  ArryResult(1,sKey+1)=nVal
End If
'ȡsKey���������
    For I=sKey+1 TO subIndex
      ArryResult(0,I+1)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,I+1)=ArryObj(1,I)
  Else
    ArryResult(1,I+1)=ArryObj(1,I)
  End If
  Next
  Else
    For I=0 TO sKey-1         'ȡsKey-1ǰ�沿������
      ArryResult(0,I)=ArryObj(0,I)
  If IsObject(ArryObj(1,I)) Then
    Set ArryResult(1,I)=ArryObj(1,I)
  Else
    ArryResult(1,I)=ArryObj(1,I)
  End If
  Next
'�����µ�����
ArryResult(0,sKey)=nKey
If IsObject(nVal) Then
  Set ArryResult(1,sKey)=nVal
Else
  ArryResult(1,sKey)=nVal
End If
'ȡsKey���������
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
ReDim ArryObj(1,CurIndex) '��������
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
CurIndex=CurIndex+1     'Insert������ָ���һ
End Sub

Public Function Exists(sKey)   '�жϴ治����ĳ���ֵ�����
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

Public Sub Remove(sKey)         '����sKey��ֵRemoveһ���ֵ�����
If Not Exists(sKey) Then
  C_ErrCode=3
  Exit Sub
End If
sType=GetType(sKey)         'ȡ��sKey�ı�������
Dim ArryResult(),I,sType,sAdd
ReDim ArryResult(1,CurIndex-2)   '����һ��������������ʱ��ŵ�
sAdd=0
If sType=0 Then             '�ַ������ͱȽ�
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
  sKey=sKey-1             '��1��Ϊ�˷����ճ�ϰ��(��1��ʼ)
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
ReDim ArryObj(1,MaxIndex) '��������
For I=0 To MaxIndex
ArryObj(0,I)=ArryResult(0,I)
If isObject(ArryResult(1,I)) Then
  Set ArryObj(1,I)=ArryResult(1,I)
Else
  ArryObj(1,I)=ArryResult(1,I)
End If
Next
Erase ArryResult
CurIndex=CurIndex-1     '��һ��Remove������ָ��
End Sub

Public Sub RemoveAll 'ȫ������ֵ�����,ֻRedimһ�¾�OK��
  Redim ArryObj(MaxIndex)
CurIndex=0
End Sub

Public Sub ClearErr   '���ô���
  C_ErrCode=0
End Sub

Private Function sIsEmpty(sVal) '�ж�sVal�Ƿ�Ϊ��ֵ
  If IsEmpty(sVal) Then
  C_ErrCode=9           'KeyֵΪ�յĴ������
  sIsEmpty=True
  Exit Function
End If
  If IsNull(sVal) Then
  C_ErrCode=9           'KeyֵΪ�յĴ������
  sIsEmpty=True
  Exit Function
End If
  If Trim(sVal)="" Then
  C_ErrCode=9           'KeyֵΪ�յĴ������
  sIsEmpty=True
  Exit Function
End If
sIsEmpty=False
End Function

Private Function GetType(sVal)   'ȡ�ñ���sVal�ı�������
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
