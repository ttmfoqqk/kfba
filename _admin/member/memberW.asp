<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim UserIdx    : UserIdx    = RequestSet("UserIdx","GET",0)
Dim pageNo     : pageNo     = RequestSet("pageNo","GET",1)
Dim sIndate    : sIndate    = RequestSet("sIndate","GET","")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate","GET","")
Dim sUserId    : sUserId    = RequestSet("sUserId","GET","")
Dim sUserName  : sUserName  = RequestSet("sUserName","GET","")
Dim sHphone3   : sHphone3   = RequestSet("sHphone3","GET","")
Dim sUserBirth : sUserBirth = RequestSet("sUserBirth","GET","")
Dim State      : State      = RequestSet("State","GET","")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sUserId="    & sUserId &_
		"&sUserName="  & sUserName &_
		"&sHphone3="   & sHphone3 &_
		"&sUserBirth=" & sUserBirth &_
		"&sState="     & sState

checkAdminLogin(g_host & g_url & "?" & PageParams  & "&UserIdx=" & UserIdx)

Call Expires()
Call dbopen()
	Call getView()

	temp_email = IIF(FV_UserEmail<>"",Split(FV_UserEmail,"@"),Split("@","@"))
	email1 = temp_email(0)
	email2 = temp_email(1)
	' 기초코드 
	Call common_code_list(10) ' 핸드폰 콤보박스 옵션
	Dim hphoneOption : hphoneOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Trim( FV_UserHphone1 ) )	
	Call common_code_list(11) ' 이메일 콤보박스 옵션	
	Dim mailOption   : mailOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Trim( email2 ) )
	Dim delfgOption  : delfgOption = "<option value=""0"" " & IIF( FV_UserDelFg="0","selected","" ) & ">사용</option><option value=""1"" " & IIF( FV_UserDelFg="1","selected","" ) & ">탈퇴</option>"
Call dbclose()

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT " &_
	"	 [UserIdx] " &_
	"	,[UserId] " &_
	"	,[UserName] " &_
	"	,[UserBirth] " &_
	"	,[UserHphone1] " &_
	"	,[UserHphone2] " &_
	"	,[UserHphone3] " &_
	"	,[UserSmsFg] " &_
	"	,[UserEmail] " &_
	"	,[UserEmailFg] " &_
	"	,[UserZipcode] " &_
	"	,[UserAddr1] " &_
	"	,[UserAddr2] " &_
	"	,[UserIndate] " &_
	"	,[UserOutdate] " &_
	"	,[UserDelFg] " &_
	"	,[UserBigo] " &_
	"FROM [dbo].[SP_USER_MEMBER] " &_
	"WHERE [UserIdx] = ?  "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput,  0, UserIdx  )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")
	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "member/leftMenu.html"
ntpl.setFile "MAIN", "member/memberW.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()


call ntpl.setBlock("MAIN", array("HIDDEN_INDATE" ,"HIDDEN_PWD_TXT" ,"HIDDEN_DELFG") )

If FV_UserIdx = "" Then 
	ntpl.tplBlockDel("HIDDEN_INDATE")
	ntpl.tplBlockDel("HIDDEN_PWD_TXT")
	ntpl.tplBlockDel("HIDDEN_DELFG")
Else
	ntpl.tplParseBlock("HIDDEN_INDATE")
	ntpl.tplParseBlock("HIDDEN_PWD_TXT")
	ntpl.tplParseBlock("HIDDEN_DELFG")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("actType", IIF( FV_UserIdx="","INSERT","UPDATE") ) _
	,array("UserIdx", FV_UserIdx ) _
	,array("UserId", FV_UserId ) _
	,array("UserName", FV_UserName ) _
	,array("UserBirth1", Mid(FV_UserBirth,1,4) ) _
	,array("UserBirth2", Mid(FV_UserBirth,5,2) ) _
	,array("UserBirth3", Mid(FV_UserBirth,7,2) ) _
	,array("UserHphone1", Trim(FV_UserHphone1) ) _
	,array("UserHphone2", Trim(FV_UserHphone2) ) _
	,array("UserHphone3", Trim(FV_UserHphone3) ) _
	,array("UserSmsFg", FV_UserSmsFg ) _
	,array("UserEmail", FV_UserEmail ) _
	,array("UserEmail1", email1 ) _
	,array("UserEmail2", email2 ) _
	,array("UserEmailFg", FV_UserEmailFg ) _
	,array("UserZipcode1", Mid(FV_UserZipcode,1,3) ) _
	,array("UserZipcode2", Mid(FV_UserZipcode,4,3) ) _
	,array("UserAddr1", FV_UserAddr1 ) _
	,array("UserAddr2", FV_UserAddr2 ) _
	,array("UserIndate", FV_UserIndate ) _
	,array("UserOutdate", FV_UserOutdate ) _
	,array("UserDelFg", FV_UserDelFg ) _
	,array("UserBigo", FV_UserBigo ) _

	,array("hphoneOption", hphoneOption ) _
	,array("mailOption", mailOption ) _
	,array("delfgOption", delfgOption ) _

	,array("pageNo"    , pageNo ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sUserId"   , sUserId ) _
	,array("sUserName" , sUserName ) _
	,array("sHphone3"  , sHphone3 ) _
	,array("sUserBirth", sUserBirth ) _
	,array("sState"    , sStateOption ) _

	,array("PageParams" , PageParams) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>