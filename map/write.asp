<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim Idx      : Idx      = RequestSet("Idx"    ,"GET",0)
Dim pageNo   : pageNo   = RequestSet("pageNo" ,"GET",1 )
Dim sPcode   : sPcode   = RequestSet("sPcode" ,"GET","")
Dim sACode   : sACode   = RequestSet("sACode" ,"GET","")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode=" & sPcode &_
		"&sACode=" & sACode

checkLogin( g_host & g_url &"?"&PageParams )

Call Expires()
Call dbopen()
	Call getData()
	actType = IIF( FI_Idx = "","INSERT", actType )

	Call common_code_list(22) ' 지역분류
	Dim addrOption  : addrOption  = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, int(IIF(FV_AddrIdx="", 0 ,FV_AddrIdx)) )

	Call common_code_list(17) ' 프로그램명
	Dim codeOption  : codeOption  = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, int(IIF(FV_CodeIdx="", IIF( sPcode="",0,sPcode ) ,FV_CodeIdx)) )


Call dbclose()


dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "map/write.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("HIDDEN_DATA_FILE"))
If Isnull(FV_Map) Or FV_Map = "" Then 
	ntpl.tplBlockDel("HIDDEN_DATA_FILE")
Else
	ntpl.tplParseBlock("HIDDEN_DATA_FILE")
End If

ntpl.tplAssign array(   _
	 array("imgDir" , TPL_DIR_IMAGES ) _
	,array("sPcode" , sPcode ) _
	,array("sACode" , sACode ) _
	,array("pageNo" , pageNo ) _
	,array("PageParams", PageParams ) _

	,array("actType", IIF( FV_Idx="","INSERT","UPDATE") ) _
	,array("Idx", FV_Idx ) _
	,array("Name", TagDecode(Trim( FV_Name )) ) _
	,array("Addr", TagDecode(Trim( FV_Addr )) ) _
	,array("Tel", TagDecode(Trim( FV_Tel )) ) _
	,array("Info", Trim( FV_Info ) ) _
	,array("WebAddr", TagDecode(Trim( FV_WebAddr )) ) _
	,array("Map", FV_Map ) _
	,array("Code", FV_Code ) _
	,array("MapImages", MapImages ) _

	,array("codeOption", codeOption ) _
	,array("addrOption", addrOption ) _

), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing





Sub getData()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Idx INT;" &_
	"SET @Idx = ?; " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,[Name] " &_
	"	,[Addr] " &_
	"	,[Tel] " &_
	"	,[Info] " &_
	"	,[WebAddr] " &_
	"	,[Map] " &_
	"	,[Code] " &_
	"	,[CodeIdx] " &_
	"	,[AddrIdx] " &_
	"FROM [dbo].[SP_PROGRAM_AREA] " &_
	"WHERE [Idx] = @Idx AND [Dellfg] = 0"

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , Idx )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")
	Set objRs = Nothing
End Sub
%>