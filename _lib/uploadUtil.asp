<%
dim resizeWidth, resizeHeight ' �̹��� ������� ���� �������� ����
' ========================================================================
' Function�� : DextFileUpload
' ��      �� : ÷������ ���ε� �� ����Ÿ ó��
' ��      �� : 
' ========================================================================
Function DextFileUpload(ByVal ControlName,ByVal sFolderName,ByVal s_w)

	set objImage = Server.CreateObject("DEXT.ImageProc") '//-- �̹��� ������� ���� �ʿ��� ��ü
	
	Dim NowYear			: NowYear		= Year(Date())
	Dim NowMonth		: NowMonth		= Month(Date())
	Dim NowDay			: NowDay		= Day(Date())
	Dim NowHour			: NowHour		= Hour(Time())
	Dim NowMinute		: NowMinute		= Minute(Time())
	Dim NowSecond		: NowSecond		= Second(Time())
	Dim NowRandomStr	: NowRandomStr	= RandomNumber(5,"")

	Dim NewFileName		: NewFileName = NowRandomStr & NowYear & NowMonth & NowDay & NowHour & NowMinute & NowSecond
	Dim sNewFileName    : sNewFileName = "s_" & NewFileName
	
	Dim f,i
	Dim arrFileName,strFilePath_new

	' �̹��� �⺻ ���� ������ 650
	Dim s_width  : s_width = 650

	If s_w > 0 Then 
		s_width = s_w
	End If
	
	Set f = UPLOAD__FORM(ControlName)
	if f <> "" then
		Dim file_ext : file_ext = mid(f.FileName, InStrRev(f.FileName, ".") + 1)	'���ϸ��� Ȯ���ڸ� �и�
		'���̸� �ߺ��˻�
		strFilePath_new = chkFileDup(sFolderName, NewFileName & "." & file_ext )
		'���� ����
		f.SaveAs strFilePath_new

		
		
		' �̹��� ���� �϶� ����� ����
		If LCase(file_ext) = "jpg" Or LCase(file_ext) = "jpeg" Or LCase(file_ext) = "gif" Or LCase(file_ext) = "bmp" Or LCase(file_ext) = "png" Then 

			'///////////////////////////////////
			'///////////////////////////////////
			'���� ������ ���ϱ�
			ImageWidth = f.ImageWidth
			ImageHeight = f.ImageHeight

			' ���ε�� �̹����� ����� ���� ũ�⺸�� Ŭ���� �����ϱ�.

			If s_width < ImageWidth Then 

				fixWidth = s_width    '## ����� ���� ������

				Call get_ImgResizeValue(ImageWidth,ImageHeight, fixWidth ) 

				If objImage.SetSourceFile(strFilePath_new) Then '-- ���ε��� ������ �����ؼ� �ִٸ�
					'jpg ����Ƽ 100%
					If LCase(file_ext) = "jpg" Or LCase(file_ext) = "jpeg" then
						 objImage.Quality = 100
					 End If
					' �������� ���ε�
					s_strFilePath_new = chkFileDup(sFolderName, sNewFileName & "." & file_ext )
					new_imagesPath = objImage.SaveAsThumbnail(s_strFilePath_new , resizeWidth, resizeHeight, false)
				end If

			End If

		End If

		arrFileName=Right(strFilePath_new, len(strFilePath_new) - instrRev(strFilePath_new, "/"))

	Else
		arrFileName=""
	End If
	DextFileUpload=arrFileName

	set objImage = Nothing ''��ü�Ҹ�
	if err <> 0 then
		alert("�����߻�")
	end if
End Function

' ==================================================================
' Function�� : ���丮 ����
' ��      �� : ���丮 ����
' ===================================================================
Sub CreateTargetFolder(strFolder)
	'dim fso
	'Set fso = Server.CreateObject("Scripting.FileSystemObject")
	'IF NOT fso.FolderExists(strFolder) Then
'		fso.CreateFolder(strFolder)
'	End IF
'	Set fso = nothing

	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	Dim sFolders : sFolders=Split(strFolder,"\")
	Dim ii
	Dim sFolderName
	
	sFolderName=sFolders(0)
	For ii=1 To UBound(sFolders)
		sFolderName=sFolderName & "\" & sFolders(ii)
		IF NOT fso.FolderExists(sFolderName) Then
			fso.CreateFolder(sFolderName)
		End IF	
	Next
	
	Set fso = nothing
End Sub

' ==================================================================
' Function�� : �ߺ��� ���ϸ� ó��
' ��      �� : �ߺ��� ���ϸ��� �ִ��� �˻��ؼ� �ٸ��̸����� ��ü
' ��      �� : FileNameWithoutExt(Ȯ���ڸ� ������ ���ϸ�), FileExt(Ȯ����)
' ��  ��  �� : chkFileDup(���ϰ�θ� ������ ���ϸ�)
' ===================================================================
Function chkFileDup(sFolderName,sFileName)
	Dim strFilePath,f_exist, count
	Dim file_ext, file_name_without_ext
	Dim FSO : Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	f_exist = true
	count = 0
	
	strFilePath=sFolderName & sFileName
	file_ext = mid(sFileName, InStrRev(sFileName, ".") + 1)				'���ϸ��� Ȯ���ڸ� �и�
	file_name_without_ext = mid(sFileName, 1, InStrRev(sFileName,".")-1)'���ϸ��� �̸��� �и�
	
	Do while f_exist
		If(fso.fileExists(strFilePath)) Then
			sFileName = file_name_without_ext & "(" & count & ")." & file_ext
			strFilePath = sFolderName & sFileName
			count = count + 1
		Else
			f_exist = false
		End If
	Loop

	chkFileDup = strFilePath
End Function
'����
Function RandomNumber(NumberLength,NumberString)
	Const DefaultString = "ABCDEFGHIJKLMNOPQRSTUVXYZ1234567890"
	Dim nCount,RanNum,nNumber,nLength

	Randomize
	If NumberString = "" Then 
		NumberString = DefaultString
	End If

	nLength = Len(NumberString)

	For nCount = 1 To NumberLength
	nNumber = Int((nLength * Rnd)+1)
	RanNum = RanNum & Mid(NumberString,nNumber,1)
	Next

	RandomNumber = RanNum
End Function

Sub get_ImgResizeValue(ByVal ImageWidth,ByVal ImageHeight, ByVal fixWidth ) 
	If ImageWidth > fixWidth then
		resizeWidth = fixWidth
		resizeHeight = ImageHeight * fixWidth / ImageWidth
	Else
		resizeWidth  = ImageWidth
		resizeHeight = ImageHeight
	End If
End Sub
%>