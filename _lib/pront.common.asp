<%
' 게시판 다운로드 패치
Dim DOWNLOAD_BASE_PATH : DOWNLOAD_BASE_PATH = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/board/&file="
' 게시판 이미지 패치
Dim BOARD_PHOTO_PATH : BOARD_PHOTO_PATH = FRONT_ROOT_DIR & "upload/board/"
' 구직게시판 다운로드 패치
Dim DOWNLOAD_BASE_PATH_JOB : DOWNLOAD_BASE_PATH_JOB = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/job/&file="
' 개인 사진 패치
Dim USER_PHOTO_PATH : USER_PHOTO_PATH = FRONT_ROOT_DIR & "upload/appMember/"
Dim DOWNLOAD_USER_PHOTO_PATH : DOWNLOAD_USER_PHOTO_PATH = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/appMember/&file="

'------------------------------------------------------------------------------------
'' 스킨 경로
'------------------------------------------------------------------------------------
Const TPL_DIR_FOLDER = "_skin/basic"
Const TPL_DIR_IMAGES = "../_skin/basic/images"
'------------------------------------------------------------------------------------
'' 사용자 로그인 체크.
'------------------------------------------------------------------------------------
Function checkLogin(url)
	If session("UserIdx")="" or IsNull(session("UserIdx"))=True Then 
		response.redirect "../member/login.asp?GoUrl=" & server.urlencode(url)
	End If
End Function

'------------------------------------------------------------------------------------
'' ntpl 로그인 블럭 처리.
'------------------------------------------------------------------------------------
Sub loginBlock_ntpl(BLOCK,LOGIN,LOGOUT)
	call ntpl.setBlock( BLOCK , array(LOGIN,LOGOUT))

	If IsNUll(Session("UserIdx")) Or Session("UserIdx")="" Then 
		ntpl.tplBlockDel(LOGOUT)
		ntpl.tplParseBlock(LOGIN)
	Else
		ntpl.tplBlockDel(LOGIN)
		ntpl.tplParseBlock(LOGOUT)
	End If
End Sub



'마이페이지 심사위원등록메뉴 카운팅
Sub CheckApplicationCnt()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT COUNT(*) AS [CNT] FROM [dbo].[SP_PROGRAM_JUDGE_APP] WHERE [UserIdx] = ? AND [Dellfg] = 0 AND [State] != 2 "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@PWD" ,adInteger , adParamInput, 0, session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "LEFT_JUDGE_MENU")
	Set objRs = Nothing
End Sub
%>