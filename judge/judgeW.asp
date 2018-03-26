<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
checkLogin( g_host & g_url)

Dim arrSubData
Dim cntSubData  : cntSubData  = -1

Dim arrJudge
Dim cntJudge    : cntJudge  = -1

Call Expires()
Call dbopen()
	Call getData()
Call dbclose()

Dim ProgramOption : ProgramOption = "<option value="""" myIdx=""0"">선택</option>"

If cntJudge > -1 Then 
	for iLoop = 0 to cntJudge
		ProgramOption = ProgramOption & "<option value="""& arrJudge(JUDGE_ProgramIdx,iLoop) &""">" & arrJudge(JUDGE_Name,iLoop) & "</option>"
	Next 
Else
	Call msgbox("등록가능한 프로그램이 없습니다.",true)
End If


dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "judge/judgeW.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("SUBDATA_LOOP" , "SUBDATA_LOOP_NODATA" ))

If cntSubData > -1 Then 

	for iLoop = 0 to cntSubData
		ntpl.setBlockReplace array( _
			 array("CompanyName", Trim( arrSubData(SUB_CompanyName,iLoop) ) )_
			,array("WorkTime", Trim( arrSubData(SUB_WorkTime,iLoop) ) )_
			,array("WorkMonth", Trim( arrSubData(SUB_WorkMonth,iLoop) ) )_
			,array("LastPosition", Trim( arrSubData(SUB_LastPosition,iLoop) ) )_
			
		), ""
		ntpl.tplParseBlock("SUBDATA_LOOP")
	Next 
	ntpl.tplBlockDel("SUBDATA_LOOP_NODATA")
Else
	ntpl.tplParseBlock("SUBDATA_LOOP_NODATA")
	ntpl.tplBlockDel("SUBDATA_LOOP")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("actType", "INSERT" ) _


	,array("ProgramOption", ProgramOption ) _
	,array("downlPhotos", DOWNLOAD_USER_PHOTO_PATH & FI_Photo ) _



	,array("UserName", FI_UserName ) _
	,array("UserId", FI_UserId ) _
	,array("UserBirth", FI_UserBirth ) _
	,array("UserPhone", FI_UserHphone1 &"-"&FI_UserHphone2&"-"&FI_UserHphone3 ) _
	,array("UserEmail", FI_UserEmail ) _
	,array("UserAddr", FI_UserAddr1 & "  " & UserAddr2 ) _
	,array("Photo", FI_Photo ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing


Sub getData()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @UserIdx INT , @PIdx INT;" &_
	"SET @UserIdx = ?; " &_
	"SET @PIdx    = ?; " &_

	"SELECT " &_
	"	 [UserName]" &_
	"	,[UserId]" &_
	"	,[UserBirth]" &_
	"	,[UserHphone1]" &_
	"	,[UserHphone2]" &_
	"	,[UserHphone3]" &_
	"	,[UserEmail]" &_
	"	,[UserAddr1]" &_
	"	,[UserAddr2]" &_
	"	,[Photo]" &_
	"FROM [dbo].[SP_USER_MEMBER] WHERE [UserIdx] = @UserIdx " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,[CompanyName] " &_
	"	,[WorkTime] " &_
	"	,[WorkMonth] " &_
	"	,[LastPosition] " &_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP_CAREER] WHERE [UserIdx] = @UserIdx " &_

	"SELECT " &_
	"	 [Idx] AS [ProgramIdx] " &_
	"	,[Name] " &_
	"FROM [dbo].[SP_COMM_CODE2] " &_
	"WHERE [PIdx] = @PIdx " &_
	"AND [UsFg] = 0 " &_
	"ORDER BY [Order] ASC , [Idx] DESC "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		.Parameters.Append .CreateParameter( "@PIdx" ,adInteger , adParamInput , 0 , 17 )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")

	'경력사항
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "SUB")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrSubData = objRs.GetRows()
		cntSubData = UBound(arrSubData, 2)
	End If

	'프로그램 옵션
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "JUDGE")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrJudge = objRs.GetRows()
		cntJudge = UBound(arrJudge, 2)
	End If

	Set objRs = Nothing
End Sub
%>