<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
checkLogin( g_host & g_url )
Call Expires()
Call dbopen()
	Call getData()
	Call common_code_list(10) ' 핸드폰 콤보박스 옵션
	Dim hphoneOption : hphoneOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Trim(FI_UserHphone1) )	
	Call common_code_list(11) ' 이메일 콤보박스 옵션	
	Dim mailOption   : mailOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Split(FI_UserEmail,"@")(1) )
	Call CheckApplicationCnt()
Call dbclose()

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "mypage/info.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA" , "MYJUDGE_LIST"))
'왼쪽 심사위원등록 메뉴
If LEFT_JUDGE_MENU_CNT > 0 Then 
	ntpl.tplParseBlock("MYJUDGE_LIST")
Else
	ntpl.tplBlockDel("MYJUDGE_LIST")
End If 

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	
	,array("actType", "UPDATE" ) _
	,array("UserIdx", Session("UserIdx") ) _
	,array("UserName", FI_UserName ) _
	,array("UserId", FI_UserId ) _
	,array("UserBirth", FI_UserBirth ) _

	,array("UserPhone2", FI_UserHphone2 ) _
	,array("UserPhone3", FI_UserHphone3 ) _
	,array("hphoneOption", hphoneOption ) _


	,array("mailOption", mailOption ) _
	,array("UserEmail1", Split(FI_UserEmail,"@")(0) ) _
	,array("UserEmail2", Split(FI_UserEmail,"@")(1) ) _

	,array("UserZip1", Mid(FI_UserZipcode,1,3) ) _
	,array("UserZip2", Mid(FI_UserZipcode,4,3) ) _
	,array("UserAddr1", FI_UserAddr1 ) _
	,array("UserAddr2", FI_UserAddr2 ) _
	,array("LastName", FI_LastName ) _
	,array("FirstName", FI_FirstName ) _
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
	"DECLARE @UserIdx INT;" &_
	"SET @UserIdx = ?; " &_

	"SELECT " &_
	"	 [UserName]" &_
	"	,[UserId]" &_
	"	,[UserBirth]" &_
	"	,[UserHphone1]" &_
	"	,[UserHphone2]" &_
	"	,[UserHphone3]" &_
	"	,[UserEmail]" &_
	"	,[UserZipcode]" &_
	"	,[UserAddr1]" &_
	"	,[UserAddr2]" &_
	"	,[Photo]" &_
	"	,[LastName]" &_
	"	,[FirstName]" &_
	"FROM [dbo].[SP_USER_MEMBER] WHERE [UserIdx] = @UserIdx "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
%>