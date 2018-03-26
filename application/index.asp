<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
'56 : 커피바리스타
'57 : 칵테일조주사
'58 : 믹솔로지스트
'59 : 와인소믈리에
'60 : 전통주관리사
'61 : 외식경영관리사
'62 : 식음료관리사


Dim applicationKey : applicationKey = RequestSet("applicationKey","GET",56)
Dim tabOnOff1 : tabOnOff1 = "_on"
Dim tabOnOff2 : tabOnOff2 = "_off"
Dim tabOnOff3 : tabOnOff3 = "_off"

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "application/application" & applicationKey & ".html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("applicationKey", applicationKey ) _
	,array("tabOnOff1", tabOnOff1 ) _
	,array("tabOnOff2", tabOnOff2 ) _
	,array("tabOnOff3", tabOnOff3 ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing
%>