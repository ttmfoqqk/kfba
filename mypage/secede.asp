<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Call Expires()
Call dbopen()
	Call CheckApplicationCnt()
Call dbclose()

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "mypage/secede.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("MYJUDGE_LIST"))
'왼쪽 심사위원등록 메뉴

If LEFT_JUDGE_MENU_CNT > 0 Then 
	ntpl.tplParseBlock("MYJUDGE_LIST")
Else
	ntpl.tplBlockDel("MYJUDGE_LIST")
End If 


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("UserId", Session("UserId") ) _
	,array("actType", "DELETE" ) _
	
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing
%>