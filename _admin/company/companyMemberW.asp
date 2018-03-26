<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
checkAdminLogin(g_host & g_url)
Dim Idx : Idx = Trim( Request.QueryString("Idx") )

Call Expires()
Call dbopen()
	If isNumeric(Idx) Then
		Call getView()

		email1 = IIF(FV_email<>"",Split(FV_email,"@")(0),"")
		email2 = IIF(FV_email<>"",Split(FV_email,"@")(1),"")
		MsgAddr1 = IIF(FV_MsgAddr<>"",Split(FV_MsgAddr,"@")(0),"")
		MsgAddr2 = IIF(FV_MsgAddr<>"",Split(FV_MsgAddr,"@")(1),"")
	End If
	' 기초코드 
	Call common_code_list(9) ' 일반전화 콤보박스 옵션
	Dim phoneOption : phoneOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Trim( FV_pHone1 ) )
	Call common_code_list(10) ' 핸드폰 콤보박스 옵션
	Dim hphoneOption : hphoneOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Trim( FV_Hphone1 ) )
	Call common_code_list(11) ' 이메일 콤보박스 옵션
	Dim mailOption : mailOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Trim( email2 ) )
	Call common_code_list(12) ' 메신저 콤보박스 옵션
	Dim msgOption : msgOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Name, CCODE_Name, Trim( MsgAddr2 ) )
Call dbclose()

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT " &_
	"	 [Idx] " &_
	"	,[Id] " &_
	"	,[Pwd] " &_
	"	,[Name] " &_
	"	,[pHone1] " &_
	"	,[pHone2] " &_
	"	,[pHone3] " &_
	"	,[Hphone1] " &_
	"	,[Hphone2] " &_
	"	,[Hphone3] " &_
	"	,[ExtNum] " &_
	"	,[DirNum] " &_
	"	,[email] " &_
	"	,[MsgAddr] " &_
	"	,[Bigo] " &_
	"	,CONVERT(VARCHAR,[Indata],23) AS [Indata] " &_
	"FROM [dbo].[SP_ADMIN_MEMBER] " &_
	"WHERE [Idx] = ?  "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput,  0, Idx  )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")
	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "company/leftMenu.html"
ntpl.setFile "MAIN", "company/companyMemberW.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("actType", IIF( FV_Idx="","INSERT","UPDATE") ) _
	,array("adminIdx", FV_Idx ) _
	,array("adminId", FV_Id ) _
	,array("adminPwd", FV_Pwd ) _
	,array("adminName", FV_Name ) _
	,array("adminPhone1", Trim(FV_pHone1) ) _
	,array("adminPhone2", Trim(FV_pHone2) ) _
	,array("adminPhone3", Trim(FV_pHone3) ) _
	,array("adminHphone1", Trim(FV_Hphone1) ) _
	,array("adminHphone2", Trim(FV_Hphone2) ) _
	,array("adminHphone3", Trim(FV_Hphone3) ) _
	,array("adminExtNum", FV_ExtNum ) _
	,array("adminDirNum", FV_DirNum ) _
	,array("adminemail", FV_email ) _
	,array("adminemail1", email1 ) _
	,array("adminemail2", email2 ) _
	,array("adminMsgAddr", FV_MsgAddr ) _
	,array("adminMsgAddr1", MsgAddr1 ) _
	,array("adminMsgAddr2", MsgAddr ) _
	,array("adminBigo", TagDecode(Trim( FV_Bigo )) ) _
	,array("phoneOption", phoneOption ) _
	,array("hphoneOption", hphoneOption ) _
	,array("mailOption", mailOption ) _
	,array("msgOption", msgOption ) _

	,array("leftMenuOverClass1"   , "" ) _
	,array("leftMenuOverClass2"   , "admin_left_over" ) _
	,array("leftMenuOverClass3"   , "" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>