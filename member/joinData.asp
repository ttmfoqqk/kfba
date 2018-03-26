<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
If Session("UserIdx") <> "" Then 
	Response.redirect "../mypage/"
End If
Dim agree      : agree      = RequestSet("agree" , "POST" , 0)
Dim authResult : authResult = RequestSet("authResult" , "POST" , "")

If agree = 0 Or authResult = "" Then 
	Call msgbox("잘못된 경로입니다.",true)
End If

'// 본인인증 세션 검사하기 //

If session("sName") = "" Then 
	Call msgbox("잘못된 경로입니다.",true)
End If

If authResult = "safe" Then 
	UserHphone1   = Mid(session("sMobileNo"),1,3)
	UserHphone2   = Mid(session("sMobileNo"),4, IIF(Len(session("sMobileNo"))<11,3,4) )
	UserHphone3   = Right(session("sMobileNo"),4)
End If


Call Expires()
Call dbopen()
	Call common_code_list(10) ' 핸드폰 콤보박스 옵션
	Dim hphoneOption : hphoneOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, UserHphone1 )	
	Call common_code_list(11) ' 이메일 콤보박스 옵션	
	Dim mailOption   : mailOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, "" )
Call dbclose()

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "member/joinData.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")
call ntpl.setBlock("MAIN", array("PHONE_AREA","PHONE_AREA_HIDDEN" ))

If authResult = "safe" Then 
	ntpl.tplParseBlock("PHONE_AREA_HIDDEN")
	ntpl.tplBlockDel("PHONE_AREA")
Else
	ntpl.tplParseBlock("PHONE_AREA")
	ntpl.tplBlockDel("PHONE_AREA_HIDDEN")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("hphoneOption", hphoneOption) _
	,array("mailOption", mailOption) _

	,array("UserPhone1", UserHphone1) _
	,array("UserPhone2", UserHphone2) _
	,array("UserPhone3", UserHphone3) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing
%>