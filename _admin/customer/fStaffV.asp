<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim Idx        : Idx      = RequestSet("Idx"    ,"GET",0)
Dim pageNo     : pageNo   = RequestSet("pageNo" ,"GET",1)
Dim actType     : actType = RequestSet("actType","GET","")

Dim sIndate    : sIndate    = RequestSet("sIndate","GET","")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate","GET","")
Dim sUserId    : sUserId    = RequestSet("sUserId","GET","")
Dim sUserName  : sUserName  = RequestSet("sUserName","GET","")
Dim sTitle     : sTitle     = RequestSet("sTitle","GET","")


Call Expires()
Call dbopen()
	Call BoardCodeList()
	Call getData()
Call dbclose()

Dim PageParams
PageParams = "pageNo="& pageNo &_
		"&sIndate="   & sIndate &_
		"&sOutdate="  & sOutdate &_
		"&sUserId="   & sUserId &_
		"&sUserName=" & sUserName &_
		"&sTitle="    & sTitle

checkAdminLogin(g_host & g_url & "?" & PageParams  & "&Idx=" & Idx)

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "customer/leftMenu.html"
ntpl.setFile "MAIN", "customer/fStaffV.html"
ntpl.setFile "FOOTER", "_inc/footer.html"
'//상단메뉴오버
Call topMenuOver()

'//왼쪽메뉴 설정
call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST"))
If BC_CNT_LIST > -1 Then 
	for iLoop = 0 to BC_CNT_LIST
		ntpl.setBlockReplace array( _
			array("Idx", BC_ARRY_LIST(BDL_Idx,iLoop) ), _
			array("Name", BC_ARRY_LIST(BDL_Name,iLoop) ) _
		), ""
		ntpl.tplParseBlock("LEFT_MENU_LIST")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST")
End If
'//왼쪽메뉴 설정 끝

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("BoardName" , BDV_Name ) _
	,array("actType"   , actType ) _
	
	,array("pageNo"    , pageNo ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sUserId"   , sUserId ) _
	,array("sUserName" , sUserName ) _
	,array("sTitle"    , sTitle ) _

	,array("PageParams" , PageParams) _

	,array("Idx"         , FI_Idx ) _
	,array("OwnerName"   , TagDecode(FI_OwnerName) ) _
	,array("ManagerName" , TagDecode(FI_ManagerName) ) _
	,array("HomePage"    , FI_HomePage ) _
	,array("CompanyName" , FI_CompanyName ) _
	,array("Addr"        , FI_Addr ) _
	,array("Tel"         , FI_Tel ) _
	,array("Fax"         , FI_Fax ) _
	,array("Email"       , FI_Email ) _
	,array("Title"       , FI_Title ) _
	,array("Form"        , FI_Form ) _
	,array("Kind"        , FI_Kind ) _
	,array("WorkArea"    , FI_WorkArea ) _
	,array("WorkTime"    , TagDecode(FI_WorkTime) ) _
	,array("StaffCnt"    , FI_StaffCnt ) _
	,array("Qualify"     , FI_Qualify ) _
	,array("Files"       , FI_Files ) _
	,array("Dates"       , FI_Dates ) _
	,array("Method"      , FI_Method ) _
	,array("Pay"         , FI_Pay ) _
	,array("insure"      , FI_insure ) _
	,array("Bigo"        , FI_Bigo ) _
	,array("InData"      , FI_InData ) _

	,array("leftMenuOverClass1"   , "admin_left_over" ) _
	,array("leftMenuOverClass2"   , "" ) _
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
	"SELECT " &_
	"	 [Idx] " &_
	"	,[OwnerName] " &_
	"	,[ManagerName] " &_
	"	,[HomePage] " &_
	"	,[CompanyName] " &_
	"	,[Addr] " &_
	"	,[Tel] " &_
	"	,[Fax] " &_
	"	,[Email] " &_
	"	,[Title] " &_
	"	,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Form] ) AS [Form] " &_
	"	,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Kind] ) AS [Kind] " &_
	"	,[WorkArea] " &_
	"	,[WorkTime] " &_
	"	,[StaffCnt] " &_
	"	,[Qualify] " &_
	"	,[Files] " &_
	"	,[Dates] " &_
	"	,[Method] " &_
	"	,[Pay] " &_
	"	,[insure] " &_
	"	,[Bigo] " &_
	"	,[InData] " &_
	"	,[UserIdx] " &_
	"	,[Pwd] " &_
	"	,[Ip] " &_
	"	,[Dellfg] " &_
	"FROM [dbo].[SP_JOB_COMPANY] " &_
	"WHERE [Idx] = ? AND [Dellfg] = 0"

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , Idx )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
%>