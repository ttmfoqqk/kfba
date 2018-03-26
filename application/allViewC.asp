<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim pageNo   : pageNo   = RequestSet("pageNo" ,"GET",1 )
Dim sPcode   : sPcode   = RequestSet("sPcode" ,"GET","")
Dim sACode   : sACode   = RequestSet("sACode" ,"GET","")

Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )
Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode=" & sPcode &_
		"&sACode=" & sACode

Dim pageUrl 
pageUrl = g_url & "?" & "pageNo=__PAGE__" &_
		"&sPcode=" & sPcode &_
		"&sACode=" & sACode

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "application/allViewC.html" ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

ntpl.tplAssign array( _
	 array("imgDir"     , TPL_DIR_IMAGES ) _
	,array("codeOption" , codeOption) _
	,array("pagelist"   , pagelist) _
	,array("pageNo"     , pageNo ) _
	,array("sPcode"     , sPcode ) _
	,array("sACode"     , sACode ) _
	,array("PageParams" , PageParams ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing
%>