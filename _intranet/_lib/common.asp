<%
' 게시판 다운로드 패치
Dim DOWNLOAD_BASE_PATH : DOWNLOAD_BASE_PATH = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/board/&file="
' 구직게시판 다운로드 패치
Dim DOWNLOAD_BASE_PATH_JOB : DOWNLOAD_BASE_PATH_JOB = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/job/&file="
' 개인 사진 패치
Dim USER_PHOTO_PATH : USER_PHOTO_PATH = FRONT_ROOT_DIR & "upload/appMember/"
Dim DOWNLOAD_USER_PHOTO_PATH : DOWNLOAD_USER_PHOTO_PATH = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/appMember/&file="
'------------------------------------------------------------------------------------
'' 스킨 경로
'------------------------------------------------------------------------------------
Const TPL_DIR_FOLDER = "_skin/basic"
Const TPL_DIR_IMAGES = "../../_admin/_skin/basic/images"
'------------------------------------------------------------------------------------
'' 관리자 로그인 체크.
'------------------------------------------------------------------------------------
Function checkAdminLogin(url)
	If session("Judge_Idx")="" or IsNull(session("Judge_Idx"))=True Then 
		response.redirect "../index.asp?GoUrl=" & server.urlencode(url)
	End If
End Function


'------------------------------------------------------------------------------------
' 왼쪽메뉴 승인된 검정에 한하여 셀렉트
'------------------------------------------------------------------------------------
Dim judge_code_arrList
Dim judge_code_cntList : judge_code_cntList = -1
Sub judge_code_list()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Judge_Idx INT,@PIdx INT;" &_
	"SET @Judge_Idx = ? " &_
	"SET @PIdx = ? " &_

	"SELECT " &_
	"	 A.[Idx] " &_
	"	,A.[Name] " &_
	"	,A.[Order] " &_
	"FROM [dbo].[SP_COMM_CODE2] A " &_
	"INNER JOIN [dbo].[SP_PROGRAM_JUDGE_APP] B ON(A.[Idx] = B.[ProgramIdx])" &_
	"WHERE [PIdx] = @PIdx " &_
	"AND [UsFg] = 0 " &_
	"AND [UserIdx] = @Judge_Idx " &_
	"AND [State] = 0 " &_
	"ORDER BY [Order] ASC , [Idx] DESC "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Judge_Idx" ,adInteger , adParamInput , 0, Session("Judge_Idx") )
		.Parameters.Append .CreateParameter( "@PIdx"      ,adInteger , adParamInput , 0, 17 )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "JCODE")
	If NOT(objRs.BOF or objRs.EOF) Then
		judge_code_arrList = objRs.GetRows()
		judge_code_cntList = UBound(judge_code_arrList, 2)
	End If
	objRs.close	: Set objRs = Nothing
End Sub
%>