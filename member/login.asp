<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
If Session("UserIdx") <> "" Then 
	Response.redirect "../mypage/"
End If
Dim GoUrl : GoUrl = RequestSet("GoUrl","GET","")

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "member/login.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("GoUrl" , GoUrl ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = nothing
%>