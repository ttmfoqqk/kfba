<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
If Session("UserIdx") <> "" Then 
	Response.redirect "../mypage/"
End If

Dim optionMonth,tmpM
For i=1 To 12
	tmpM = IIF( i < 10 , "0" & i , i )
	optionMonth = optionMonth & "<option value='" & tmpM & "'>" & tmpM & "</option>"
Next

Dim optionDay,tmpD
For i=1 To 31
	tmpD = IIF( i < 10 , "0" & i , i )
	optionDay = optionDay & "<option value='" & tmpD & "'>" & tmpD & "</option>"
Next

Call Expires()
Call dbopen()
	Call common_code_list(10) ' 핸드폰 콤보박스 옵션
	Dim hphoneOption : hphoneOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, "" )	
	Call common_code_list(11) ' 이메일 콤보박스 옵션	
	Dim mailOption   : mailOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, "" )
Call dbclose()

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "member/fId.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("optionMonth", optionMonth ) _
	,array("optionDay", optionDay ) _
	,array("hphoneOption", hphoneOption ) _
	,array("mailOption", mailOption ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing
%>