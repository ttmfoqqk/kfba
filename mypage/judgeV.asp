<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim arrSubData
Dim cntSubData  : cntSubData  = -1

Dim Idx      : Idx = RequestSet("Idx","GET",0)
Dim pageNo   : pageNo   = RequestSet("pageNo"   ,"GET",1 )

Dim PageParams : PageParams = "pageNo=" & pageNo

checkLogin( g_host & g_url &"?"&PageParams )

Call Expires()
Call dbopen()
	Call getData()
	Call CheckApplicationCnt()
Call dbclose()

If FI_State = "0" Then 
	StateTxt = "<font color=""#11179a"">승인</font>"
ElseIf FI_State = "1" Then 
	StateTxt = "접수"
ElseIf FI_State = "2" Then 
	StateTxt = "<font color=""#9a1134"">불합격</font>"
End If

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "mypage/judgeV.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("SUBDATA_LOOP" , "SUBDATA_LOOP_NODATA" , "MYJUDGE_LIST" ))

'왼쪽 심사위원등록 메뉴
If LEFT_JUDGE_MENU_CNT > 0 Then 
	ntpl.tplParseBlock("MYJUDGE_LIST")
Else
	ntpl.tplBlockDel("MYJUDGE_LIST")
End If 


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
	,array("actType", "MODIFY" ) _
	,array("pageNo", pageNo ) _
	,array("Idx"    , Idx ) _


	,array("ProgramName", FI_ProgramName & IIF( FI_ProgramKind="1"," [필기]"," [실기]" ) ) _
	,array("State", StateTxt ) _
	,array("downlPhotos", DOWNLOAD_USER_PHOTO_PATH & FI_Photo ) _
	,array("downlFile", DOWNLOAD_USER_PHOTO_PATH & FI_FileName ) _



	,array("UserName", FI_UserName ) _
	,array("UserId", FI_UserId ) _
	,array("UserBirth", FI_UserBirth ) _
	,array("UserPhone", FI_UserHphone1 &"-"&FI_UserHphone2&"-"&FI_UserHphone3 ) _
	,array("UserEmail", FI_UserEmail ) _
	,array("UserAddr", FI_UserAddr1 & "  " & UserAddr2 ) _
	,array("Photo", FI_Photo ) _
	,array("FileName", FI_FileName ) _
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
	"DECLARE @UserIdx INT , @Idx INT;" &_
	"SET @UserIdx = ?; " &_
	"SET @Idx     = ?; " &_

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
	"	,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = B.[ProgramIdx] ) AS [ProgramName] "&_
	"	,[State] "&_
	"	,[FileName] " &_
	"	,[ProgramKind] " &_
	"FROM [dbo].[SP_USER_MEMBER] A " &_
	"INNER JOIN [dbo].[SP_PROGRAM_JUDGE_APP] B ON(A.[UserIdx] = B.[UserIdx] )" &_
	"WHERE A.[UserIdx] = @UserIdx " &_
	"AND [Idx] = @Idx " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,[CompanyName] " &_
	"	,[WorkTime] " &_
	"	,[WorkMonth] " &_
	"	,[LastPosition] " &_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP_CAREER] WHERE [UserIdx] = @UserIdx "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		.Parameters.Append .CreateParameter( "@Idx"     ,adInteger , adParamInput , 0 , Idx )
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

	Set objRs = Nothing
End Sub
%>