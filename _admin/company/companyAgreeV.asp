<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
checkAdminLogin( g_host & g_url )

Call Expires()
Call dbopen()
	Call getData()
Call getData()

Sub getData()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT [Agree1] , [Agree2] "  &_
	" FROM [dbo].[SP_COMM_AGREE]"

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

Dim GoUrl : GoUrl = Request.QueryString("GoUrl")

'// ���ø� ���丮 ���� (�⺻ tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "company/leftMenu.html"
ntpl.setFile "MAIN", "company/companyAgreeV.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//��ܸ޴�����
Call topMenuOver()

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("agree1", FI_Agree1 ) _
	,array("agree2", FI_Agree2 ) _

	,array("leftMenuOverClass1"   , "admin_left_over" ) _
	,array("leftMenuOverClass2"   , "" ) _
	,array("leftMenuOverClass3"   , "" ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = nothing

%>