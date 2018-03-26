<%
dim resizeWidth, resizeHeight ' 이미지 리사이즈를 위해 전역변수 설정
' ========================================================================
' Function명 : DextFileUpload
' 목      적 : 첨부파일 업로드 및 데이타 처리
' 입      력 : 
' ========================================================================
Function DextFileUpload(ByVal ControlName,ByVal sFolderName,ByVal s_w)

	set objImage = Server.CreateObject("DEXT.ImageProc") '//-- 이미지 리사이즈를 위해 필요한 객체
	
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

	' 이미지 기본 가로 사이즈 650
	Dim s_width  : s_width = 650

	If s_w > 0 Then 
		s_width = s_w
	End If
	
	Set f = UPLOAD__FORM(ControlName)
	if f <> "" then
		Dim file_ext : file_ext = mid(f.FileName, InStrRev(f.FileName, ".") + 1)	'파일명에서 확장자만 분리
		'새이름 중복검사
		strFilePath_new = chkFileDup(sFolderName, NewFileName & "." & file_ext )
		'원본 저장
		f.SaveAs strFilePath_new

		
		
		' 이미지 파일 일때 썸네일 저장
		If LCase(file_ext) = "jpg" Or LCase(file_ext) = "jpeg" Or LCase(file_ext) = "gif" Or LCase(file_ext) = "bmp" Or LCase(file_ext) = "png" Then 

			'///////////////////////////////////
			'///////////////////////////////////
			'원본 사이즈 구하기
			ImageWidth = f.ImageWidth
			ImageHeight = f.ImageHeight

			' 업로드된 이미지가 썸네일 설정 크기보다 클때만 생성하기.

			If s_width < ImageWidth Then 

				fixWidth = s_width    '## 썸네일 가로 사이즈

				Call get_ImgResizeValue(ImageWidth,ImageHeight, fixWidth ) 

				If objImage.SetSourceFile(strFilePath_new) Then '-- 업로드한 파일을 지정해서 있다면
					'jpg 퀄리티 100%
					If LCase(file_ext) = "jpg" Or LCase(file_ext) = "jpeg" then
						 objImage.Quality = 100
					 End If
					' 리사이즈 업로드
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

	set objImage = Nothing ''객체소멸
	if err <> 0 then
		alert("에러발생")
	end if
End Function

' ==================================================================
' Function명 : 디렉토리 생성
' 목      적 : 디렉토리 생성
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
' Function명 : 중복된 파일명 처리
' 목      적 : 중복된 파일명이 있는지 검사해서 다른이름으로 대체
' 입      력 : FileNameWithoutExt(확장자를 제외한 파일명), FileExt(확장자)
' 리  턴  값 : chkFileDup(파일경로를 포함한 파일명)
' ===================================================================
Function chkFileDup(sFolderName,sFileName)
	Dim strFilePath,f_exist, count
	Dim file_ext, file_name_without_ext
	Dim FSO : Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	f_exist = true
	count = 0
	
	strFilePath=sFolderName & sFileName
	file_ext = mid(sFileName, InStrRev(sFileName, ".") + 1)				'파일명에서 확장자만 분리
	file_name_without_ext = mid(sFileName, 1, InStrRev(sFileName,".")-1)'파일명에서 이름만 분리
	
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
'랜덤
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