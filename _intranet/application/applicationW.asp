<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim Idx        : Idx        = RequestSet("Idx"      , "GET" , 0)
Dim pageNo     : pageNo     = RequestSet("pageNo"   , "GET" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate"  , "GET" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate" , "GET" , "")
Dim sOnDate    : sOnDate    = RequestSet("sOnDate"  , "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "56")
Dim sArea      : sArea      = RequestSet("sArea"    , "GET" , "")

Dim sId        : sId        = RequestSet("sId"      , "GET" , "")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"  , "GET" , "")
Dim sState     : sState     = RequestSet("sState"   , "GET" , "")
Dim sSnumber   : sSnumber   = RequestSet("sSnumber" , "GET" , "")
Dim sKind      : sKind      = RequestSet("sKind"    , "GET" , "")
Dim sClass     : sClass     = RequestSet("sClass"   , "GET" , "")

Dim sOnTime    : sOnTime    = RequestSet("sOnTime"  , "GET" , "")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sArea="      & sArea &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sState="     & sState &_
		"&sSnumber="   & sSnumber &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass &_
		"&sOnTime="    & sOnTime

checkAdminLogin(g_host & g_url  & "?" & PageParams & "&Idx=" & Idx)



Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )
	Call getView()
Call dbclose()

PhotoExt = FILE_CHECK_EXT_RETURN( FI_Photo )
If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
	UserPhotos = img_resize(USER_PHOTO_PATH, FI_Photo ,150,200)
Else
	UserPhotos= "<a href="""&DOWNLOAD_FI_Photo_PATH &  FI_Photo &""">"& FI_Photo &"</a>"
End If

Dim StateOption

StateOption = "" &_
"<option value=""1"" "&IIF(FV_State="1","selected","")&" >입금대기</option>" &_
"<option value=""0"" "&IIF(FV_State="0","selected","")&" >접수완료</option>" &_
"<option value=""2"" "&IIF(FV_State="2","selected","")&" >접수취소</option>" &_
"<option value=""3"" "&IIF(FV_State="3","selected","")&" >불합격</option>" &_
"<option value=""4"" "&IIF(FV_State="4","selected","")&" >미응시(불합격)</option>" &_
"<option value=""10"" "&IIF(FV_State="10","selected","")&" >합격</option>"


StateTxt = ""
If FV_State = "0" Then 
	StateTxt = "접수완료"
ElseIf FV_State = "1" Then 
	StateTxt = "<font color=""#11179a"">입금대기</font>"
ElseIf FV_State = "2" Then 
	StateTxt = "<font color=""#9a1134"">접수취소</font>"
ElseIf FV_State = "3" Then 
	StateTxt = "<font color=""#9a1134"">불합격</font>"
ElseIf FV_State = "4" Then 
	StateTxt = "<font color=""#9a1134"">미응시(불합격)</font>"
ElseIf FV_State = "10" Then 
	StateTxt = "<font color=""#11179a"">합격</font>"
End If


Dim PayModeTxt
If FV_PayMode = "SC0010" Then 
	PayModeTxt = "카드결제"
ElseIf FV_PayMode = "SC0030" Then 
	PayModeTxt = "은행결제"
ElseIf FV_PayMode = "SC0060" Then 
	PayModeTxt = "핸드폰결제"
ElseIf FV_PayMode = "SC0040" Then 
	PayModeTxt = "무통장입금"
End If

PrograName = FV_ProgramName

If FV_Class = "1" Then
	PrograName = PrograName & " 1급"
ElseIf FV_Class = "2" Then
	PrograName = PrograName & " 2급"
ElseIf FV_Class = "3" Then
	PrograName = PrograName & " [SPECIAL]"
End If

If FV_Kind = "1" Then
	PrograName = PrograName & " [필기]"
ElseIf FV_Kind = "2" Then
	PrograName = PrograName & " [실기]"
End If

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @UserIdx INT,@Idx INT,@Judge_Idx INT;" &_
	"SET @Idx        = ? ;" &_
	"SET @Judge_Idx  = ? ;" &_
	"SET @UserIdx    = ( " &_
	"	SELECT A.[UserIdx] " &_
	"	FROM [dbo].[SP_PROGRAM_APP] A " &_
	"	INNER JOIN [dbo].[SP_PROGRAM] B ON(A.[ProgramIdx] = B.[Idx]) " &_
	"	WHERE A.[Idx] = @Idx " &_
	"	) ; " &_
	"SELECT " &_
	"	 A.[Idx] " &_
	"	,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = B.[CodeIdx] ) AS [ProgramName] " &_
	"	,A.[InData] " &_
	"	,A.[State] " &_
	"	,A.[PayMode]" &_
	"	,A.[Bigo] " &_
	"	,A.[NocachDate] " &_
	"	,B.[OnData] " &_
	"	,B.[Kind] " &_
	"	,B.[Class] " &_
	"	,ISNULL( B.[Pay],0 ) AS [Pay] " &_
	"	,C.[Name] AS [AreaName] " &_
	"FROM [dbo].[SP_PROGRAM_APP] A " &_
	"INNER JOIN [dbo].[SP_PROGRAM] B ON(A.[ProgramIdx] = B.[Idx]) " &_
	"INNER JOIN [dbo].[SP_PROGRAM_AREA] C ON(A.[AreaIdx] = C.[Idx]) " &_
	"WHERE A.[Idx] = @Idx " &_


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
	"	,[LastName]" &_
	"	,[FirstName]" &_
	"FROM [dbo].[SP_USER_MEMBER] WHERE [UserIdx] = @UserIdx "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"       ,adInteger , adParamInput , 0 , Idx )
		.Parameters.Append .CreateParameter( "@Judge_Idx" ,adInteger , adParamInput , 0 , Session("Judge_Idx") )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")

	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "FI")
	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( INTRA_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "application/leftMenu.html"
ntpl.setFile "MAIN", "application/applicationW.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(common_code_arrList(CCODE_Idx,iLoop))=sPcode,"admin_left_over","" ) ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST")
End If


ntpl.tplAssign array(   _
	 array("imgDir"      , TPL_DIR_IMAGES ) _
	,array("pageList"    , pagelist ) _
	,array("PageParams"  , PageParams ) _
	,array("pageNo"      , pageNo ) _
	,array("sIndate"     , sIndate ) _
	,array("sOutdate"    , sOutdate ) _
	,array("sOnDate"     , sOnDate ) _
	,array("sPcode"      , sPcode ) _
	,array("sArea"       , sArea ) _
	,array("sId"         , sId ) _
	,array("sName"       , sName ) _
	,array("sPhone3"     , sPhone3 ) _
	,array("sState"      , sState ) _
	,array("sSnumber"    , sSnumber ) _
	,array("sKind"       , sKind ) _
	,array("sClass"      , sClass ) _
	,array("sOnTime"     , sOnTime ) _
	,array("actType"     , IIF( FV_Idx="","INSERT","UPDATE") ) _
	,array("Idx"         , FV_Idx ) _
	,array("ProgramName" , PrograName ) _
	,array("OnData"      , FV_OnData ) _
	,array("InData"      , FV_InData ) _
	,array("AreaName"    , TagDecode(Trim( FV_AreaName )) ) _
	,array("State"       , FV_State ) _
	,array("StateOption" , StateOption ) _
	,array("StateTxt"    , StateTxt ) _	
	,array("Pay"         , FormatNumber(FV_Pay,0) ) _
	,array("PayMode"     , PayModeTxt ) _
	,array("PayDate"     , IIF( FV_PayMode = "SC0040" , IIF( FV_NocachDate="","미입금", FV_NocachDate ) , FV_InData ) ) _
	,array("Bigo"        , TagDecode(FV_Bigo) ) _
	,array("Snumber"     , IIF( FV_Snumber="","접수완료 후 수검번호가 부여됩니다.",FV_Snumber ) ) _
	,array("UserIdx"     , Session("UserIdx") ) _
	,array("UserName"    , FI_UserName ) _
	,array("UserId"      , FI_UserId ) _
	,array("UserBirth"   , FI_UserBirth ) _
	,array("UserPhone"   , FI_UserHphone1 &"-"&FI_UserHphone2&"-"&FI_UserHphone3 ) _
	,array("UserEmail"   , FI_UserEmail ) _
	,array("UserAddr"    , FI_UserAddr1 & "  " & UserAddr2 ) _
	,array("LastName"    , IIF(FI_LastName="","&nbsp;",FI_LastName) ) _
	,array("FirstName"   , IIF(FI_FirstName="","&nbsp;",FI_FirstName) ) _
	,array("Photo"       , UserPhotos ) _
	,array("leftMenuOverClass1" , "admin_left_over" ) _
	,array("leftMenuOverClass2" , "" ) _

	,array("Judge_Id"    , Session("Judge_Id") ) _
	,array("Judge_Name"  , Session("Judge_Name") ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>