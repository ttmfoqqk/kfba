<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
'56 : Ŀ�ǹٸ���Ÿ
'57 : Ĭ�������ֻ�
'58 : �ͼַ�����Ʈ
'59 : ���μҹɸ���
'60 : �����ְ�����
'61 : �ܽİ濵������
'62 : �����������


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
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("applicationKey", applicationKey ) _
	,array("tabOnOff1", tabOnOff1 ) _
	,array("tabOnOff2", tabOnOff2 ) _
	,array("tabOnOff3", tabOnOff3 ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = nothing
%>