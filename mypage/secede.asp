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
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("MYJUDGE_LIST"))
'���� �ɻ�������� �޴�

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

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = nothing
%>