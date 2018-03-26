<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%

Dim arrSubData
Dim cntSubData : cntSubData  = -1

Dim Idx        : Idx        = RequestSet("Idx" , "GET" , 0)
Dim pageNo     : pageNo     = RequestSet("pageNo" , "GET" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate" , "GET" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate", "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "56")
Dim sState     : sState     = RequestSet("sState"    , "GET" , "")
Dim sId        : sId        = RequestSet("sId"    , "GET" , "")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"    , "GET" , "")
Dim sBirth     : sBirth     = RequestSet("sBirth"    , "GET" , "")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sPcode="     & sPcode &_
		"&sState="     & sState &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sBirth="     & sBirth

checkAdminLogin(g_host & g_url  & "?" & PageParams  & "&Idx=" & Idx)

Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Call getView()
Call dbclose()

PrograName = FI_ProgramName

If FI_ProgramKind = "1" Then
	PrograName = PrograName & " [필기]"
ElseIf FI_ProgramKind = "2" Then
	PrograName = PrograName & " [실기]"
ElseIf FI_ProgramKind = "3" Then
	PrograName = PrograName & " [SPECIAL]"
End If

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @UserIdx INT,@Idx INT;" &_
	"SET @Idx = ?; " &_
	"SET @UserIdx = (SELECT [UserIdx] FROM [dbo].[SP_PROGRAM_JUDGE_APP] WHERE [Idx] = @Idx ) ; " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [ProgramIdx] ) AS [ProgramName] " &_
	"	,[UserIdx] " &_
	"	,[State] " &_
	"	,[InData] " &_
	"	,[Ip] " &_
	"	,[Dellfg] " &_
	"	,[ProgramKind] " &_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP] WHERE [Idx] = @Idx " &_

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
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP_CAREER] WHERE [UserIdx] = @UserIdx "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput, 0, Idx )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")

	'유저정보
	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "USER")

	'경력사항
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "SUB")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrSubData = objRs.GetRows()
		cntSubData = UBound(arrSubData, 2)
	End If
	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "application/leftMenu.html"
ntpl.setFile "MAIN", "application/judgeW.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()

call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST1","LEFT_MENU_LIST2"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST1")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST1")
End If

If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(common_code_arrList(CCODE_Idx,iLoop))=sPcode,"admin_left_over","" ) ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST2")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST2")
End If

call ntpl.setBlock("MAIN", array("SUBDATA_LOOP" ))

If cntSubData > -1 Then 

	for iLoop = 0 to cntSubData
		ntpl.setBlockReplace array( _
			 array("CompanyName", arrSubData(SUB_CompanyName,iLoop) )_
			,array("WorkTime", arrSubData(SUB_WorkTime,iLoop) )_
			,array("WorkMonth", arrSubData(SUB_WorkMonth,iLoop) )_
			,array("LastPosition", arrSubData(SUB_LastPosition,iLoop) )_
			
		), ""
		ntpl.tplParseBlock("SUBDATA_LOOP")
	Next 
Else
	ntpl.tplBlockDel("SUBDATA_LOOP")
End If

PhotoExt = FILE_CHECK_EXT_RETURN( USER_Photo )
If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
	UserPhotos = img_resize(USER_PHOTO_PATH, USER_Photo ,150,200)
Else
	UserPhotos= "<a href="""&DOWNLOAD_USER_Photo_PATH &  USER_Photo &""">"& USER_Photo &"</a>"
End If

StateOption = "" &_
"<option value=""1"" "&IIF(FI_State="1","selected","")&" >접수</option>" &_
"<option value=""0"" "&IIF(FI_State="0","selected","")&" >승인</option>" &_
"<option value=""2"" "&IIF(FI_State="2","selected","")&" >불합격</option>" 


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("pageList", pagelist ) _
	,array("PageParams", PageParams ) _

	,array("pageNo"   , pageNo ) _
	,array("sIndate"  , sIndate ) _
	,array("sOutdate" , sOutdate ) _
	,array("sPcode"   , sPcode ) _
	,array("sState"   , sState ) _
	,array("sId"      , sId ) _
	,array("sName"    , sName ) _
	,array("sPhone3"  , sPhone3 ) _
	,array("sBirth"   , sBirth ) _

	,array("actType", IIF( FI_Idx="","INSERT","UPDATE") ) _
	,array("Idx", FI_Idx ) _
	,array("UserId", USER_UserId ) _
	,array("UserName", USER_UserName ) _
	,array("UserBirth", USER_UserBirth ) _
	,array("UserEmail", USER_UserEmail ) _
	,array("UserPhone", USER_UserHphone1 &"-"& USER_UserHphone2 &"-"& USER_UserHphone3 ) _
	,array("UserAddr", USER_UserAddr1 &" "& USER_UserAddr2 ) _
	,array("ProgramName", PrograName ) _
	,array("Photo", UserPhotos ) _
	,array("StateOption", StateOption ) _

	,array("leftMenuOverClass1"   , "" ) _
	,array("leftMenuOverClass2"   , "admin_left_over" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>