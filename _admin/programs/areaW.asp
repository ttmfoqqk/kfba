<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim Idx      : Idx    = RequestSet("Idx" , "GET" , 0)
Dim pageNo   : pageNo = RequestSet("pageNo" , "GET" , 1)

Dim sIndate  : sIndate   = RequestSet("sIndate" , "GET" , "")
Dim sOutdate : sOutdate  = RequestSet("sOutdate", "GET" , "")
Dim sName    : sName     = RequestSet("sName"   , "GET" , "")
Dim sAddr    : sAddr     = RequestSet("sAddr"   , "GET" , "")
Dim sCode    : sCode     = RequestSet("sCode"   , "GET" , "")
Dim sTel     : sTel      = RequestSet("sTel"    , "GET" , "")
Dim sPcode   : sPcode    = RequestSet("sPcode"  , "GET" , "56")
Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode="     & sPcode &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sName="      & sName &_
		"&sAddr="      & sAddr &_
		"&sCode="      & sCode &_
		"&sTel="       & sTel

checkAdminLogin(g_host & g_url & "?" & PageParams & "&Idx=" & Idx)


Call Expires()
Call dbopen()
	Call getView()

	Call common_code_list(22) ' 지역분류
	Dim addrOption  : addrOption  = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, int(IIF(FV_AddrIdx="", 0 ,FV_AddrIdx)) )

	Call common_code_list(17) ' 프로그램명
	Dim codeOption  : codeOption  = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, int(IIF(FV_CodeIdx="", IIF( sPcode="",0,sPcode ) ,FV_CodeIdx)) )	
	
Call dbclose()

PhotoExt = FILE_CHECK_EXT_RETURN( FV_Map )
If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
	MapImages = img_resize( "/upload/programsArea/", FV_Map ,300,300)
Else
	MapImages= "<a href=""_lib/dowload.asp?pach=/upload/programsArea/&file="&FV_Map&""">"& FV_Map &"</a>"
End If

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT " &_
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
	"	,[IntranetPwd] " &_
	"FROM [dbo].[SP_PROGRAM_AREA] " &_
	"WHERE [Idx] = ? AND [Dellfg] = 0 "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput, 0, Idx )
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
ntpl.setFile "LEFT", "programs/leftMenu.html"
ntpl.setFile "MAIN", "programs/areaW.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()

call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST1","LEFT_MENU_LIST2"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("sKey", common_code_arrList(CCODE_sKey,iLoop) ) _
			
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST1")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST1")
End If

If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(common_code_arrList(CCODE_Idx,iLoop))=sPcode,"admin_left_over","" ) ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST2")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST2")
End If

call ntpl.setBlock("MAIN", array("HIDDEN_DATA_FILE"))
If Isnull(FV_Map) Or FV_Map = "" Then 
	ntpl.tplBlockDel("HIDDEN_DATA_FILE")
Else
	ntpl.tplParseBlock("HIDDEN_DATA_FILE")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("pageList", pagelist ) _
	,array("PageParams", PageParams ) _

	,array("pageNo"    , pageNo ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sName", sName ) _
	,array("sAddr", sAddr ) _
	,array("sCode" , sCode  ) _
	,array("sTel" , sTel  ) _
	,array("sPcode", sPcode ) _

	,array("actType", IIF( FV_Idx="","INSERT","UPDATE") ) _
	,array("Idx", FV_Idx ) _
	,array("Name", Trim( FV_Name ) ) _
	,array("Addr", TagDecode(Trim( FV_Addr )) ) _
	,array("Tel", TagDecode(Trim( FV_Tel )) ) _
	,array("Info", Trim( FV_Info ) ) _
	,array("WebAddr", TagDecode(Trim( FV_WebAddr )) ) _
	,array("Map", FV_Map ) _
	,array("Code", FV_Code ) _
	,array("MapImages", MapImages ) _
	,array("IntranetPwd", TagDecode(FV_IntranetPwd) ) _
	

	,array("codeOption", codeOption ) _
	,array("addrOption", addrOption ) _

	,array("leftMenuOverClass1"   , "" ) _
	,array("leftMenuOverClass2"   , "admin_left_over" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>