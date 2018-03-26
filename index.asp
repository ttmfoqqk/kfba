<!-- #include file = "./_lib/header.asp" -->
<!-- #include file = "./_lib/template.class.asp" -->
<!-- #include file = "_lib/pront.common.asp" -->
<%
Dim arrAppl    , arrNoti  , arrgallery
Dim cntAppl    : cntAppl    = -1
Dim cntNoti    : cntNoti    = -1
Dim cntgallery : cntgallery = -1
Dim boardKey   : boardKey   = 1  ' 공지사항 IDX
Dim galleryKey : galleryKey = 14 ' 갤러리 IDX

Call Expires()
Call dbopen()
	Call getList()
Call dbclose()


dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "main/main.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""

'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")
call ntpl.setBlock("MAIN", array("BOARD_LOOP","GALLERY_LOOP" ,"APPL_LOOP"))

call ntpl.setBlock("FOOTER", array("QUICK_MENU"))
ntpl.tplBlockDel("QUICK_MENU")

'공지사항
If cntNoti > -1 Then 
	for iLoop = 0 to cntNoti
		ntpl.setBlockReplace array( _
			 array("Idx", arrNoti(NT_Idx,iLoop) ) _
			,array("Title", HtmlTagRemover( arrNoti(NT_Title,iLoop) , 38 ) ) _
			,array("Indate", arrNoti(NT_Indate,iLoop) ) _
		), ""
		ntpl.tplParseBlock("BOARD_LOOP")
	Next 
Else
	ntpl.tplBlockDel("BOARD_LOOP")
End If
'갤러리
If cntgallery > -1 Then 
	for iLoop = 0 to cntgallery

		PhotoExt = FILE_CHECK_EXT_RETURN( arrgallery(GL_File,iLoop) )
		If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
			Photos = "<div style=""padding-bottom:10px;"">" & img_resize(BOARD_PHOTO_PATH, arrgallery(GL_File,iLoop) ,93,61) & "</div>"
		End If

		ntpl.setBlockReplace array( _
			 array("Idx"  , arrgallery(GL_Idx,iLoop) ) _
			,array("Title", HtmlTagRemover( arrgallery(GL_Title,iLoop) , 13 ) ) _
			,array("Photo", Photos ) _
		), ""
		ntpl.tplParseBlock("GALLERY_LOOP")
	Next 
Else
	ntpl.tplBlockDel("GALLERY_LOOP")
End If
'응시일정
If cntAppl > -1 Then 
	for iLoop = 0 to cntAppl
		
		PrograName = arrAppl(AP_ProgramName,iLoop)

		If arrAppl(AP_Class,iLoop) = "1" Then
			PrograName = PrograName & " 1급"
		ElseIf arrAppl(AP_Class,iLoop) = "2" Then
			PrograName = PrograName & " 2급"
		End If

		If arrAppl(AP_Kind,iLoop) = "1" Then
			PrograName = PrograName & " [필기]"
		ElseIf arrAppl(AP_Kind,iLoop) = "2" Then
			PrograName = PrograName & " [실기]"
		End If

		ProgramLink = "./application/write.asp?applicationKey=" & arrAppl(AP_CodeIdx,iLoop)
		' 마감
		If arrAppl(AP_EndDate,iLoop) < Left(Now(),10) Then
			ProgramLink = "javascript:void(alert('응시 마감되었습니다.'))"
		End If
		' 접수전
		If arrAppl(AP_StartDate,iLoop) > Left(Now(),10) Then 
			ProgramLink = "javascript:void(alert('응시 접수기간이 아닙니다.'))"
		End If
		' 인원제한
		If arrAppl(AP_MaxNumber,iLoop) <= arrAppl(AP_CNT_APP,iLoop) Then 
			ProgramLink = "javascript:void(alert('응시 정원초과!'))"
		End If

		ntpl.setBlockReplace array( _
			 array("CodeIdx"    , arrAppl(AP_CodeIdx,iLoop) ) _
			,array("Link"       , ProgramLink ) _
			,array("OnData"     , arrAppl(AP_OnData,iLoop) ) _
			,array("ProgramName", PrograName ) _
		), ""
		ntpl.tplParseBlock("APPL_LOOP")
	Next 
Else
	ntpl.tplBlockDel("APPL_LOOP")
End If


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _

	,array("boardKey", boardKey) _
	,array("galleryKey", galleryKey) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing


Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @BoardKey INT,@galleryKey INT;" &_
	"SET @BoardKey   = ?; " &_
	"SET @galleryKey = ?; " &_

	"SELECT TOP 5 " &_
	"	 [Idx] " &_
	"	,[Title] " &_
	"	,CONVERT(VARCHAR(10),[Indate],111) AS [Indate] " &_
	"FROM [dbo].[SP_BOARD] " &_
	"WHERE [BoardKey] = @BoardKey AND [Dellfg] = 0 " &_
	"ORDER BY [Notice] DESC , [Idx] DESC; " &_

	"SELECT TOP 4 " &_
	"	 [Idx] " &_
	"	,[Title] " &_
	"	,[File] " &_
	"	,CONVERT(VARCHAR(10),[Indate],111) AS [Indate] " &_
	"FROM [dbo].[SP_BOARD] " &_
	"WHERE [BoardKey] = @galleryKey AND [Dellfg] = 0 AND [File] <> '' AND [File] is not null " &_
	"ORDER BY [Idx] DESC; " &_

	"SELECT TOP 5 " &_
	"	 A.[Idx] " &_
	"	,A.[CodeIdx] " &_
	"	,CONVERT(varchar(10),A.[StartDate],23) AS [StartDate] " &_
	"	,CONVERT(varchar(10),A.[EndDate],23) AS [EndDate] " &_
	"	,ISNULL( A.[MaxNumber] , 0 ) AS [MaxNumber] " &_
	"	,A.[Kind] " &_
	"	,A.[Class] " &_
	"	,CONVERT(VARCHAR(10),A.[OnData],111) AS [OnData] " &_
	"	,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = A.[CodeIdx] ) AS [ProgramName] " &_
	"FROM [dbo].[SP_PROGRAM] A " &_
	"LEFT JOIN ( " &_
	"	SELECT " &_
	"		 [ProgramIdx] " &_
	"		,COUNT(*) AS [CNT_APP] " &_
	"	FROM [dbo].[SP_PROGRAM_APP] " &_
	"	WHERE [State] != 2 " &_
	"	GROUP BY [ProgramIdx] " &_
	") B ON(A.[Idx] = B.[ProgramIdx] ) " &_
	"WHERE [Dellfg] = 0 " &_
	"ORDER BY A.[OnData] DESC; "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@BoardKey"   ,adInteger , adParamInput ,  0  , BoardKey )
		.Parameters.Append .CreateParameter( "@galleryKey" ,adInteger , adParamInput ,  0  , galleryKey )

		set objRs = .Execute
	End with
	call cmdclose()
	'공지사항 리스트
	CALL setFieldIndex(objRs, "NT")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrNoti = objRs.GetRows()
		cntNoti = UBound(arrNoti, 2)
	End If
	'갤러리
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "GL")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrgallery = objRs.GetRows()
		cntgallery = UBound(arrgallery, 2)
	End If
	'응시일정
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "AP")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrAppl = objRs.GetRows()
		cntAppl = UBound(arrAppl, 2)
	End If
	Set objRs = Nothing
End Sub
%>