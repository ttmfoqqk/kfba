<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim Idx      : Idx      = RequestSet("Idx"    ,"GET",0)
Dim pageNo   : pageNo   = RequestSet("pageNo" ,"GET",1)

Dim sName    : sName    = RequestSet("sName"    ,"GET",0)
Dim sId      : sId      = RequestSet("sId"      ,"GET",0)
Dim sTitle   : sTitle   = RequestSet("sTitle"   ,"GET",0)
Dim sContant : sContant = RequestSet("sContant" ,"GET",0)
Dim sWord    : sWord    = RequestSet("sWord"    ,"GET","")


Call Expires()
Call dbopen()
	Call getData()
Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&BoardKey=" & BoardKey &_
		"&sName="    & sName &_
		"&sId="      & sId &_
		"&sTitle="   & sTitle &_
		"&sContant=" & sContant &_
		"&sWord="    & sWord

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "comunity/fStaffV.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("BOARD_BTN_WRITE"))
'ntpl.tplBlockDel("BOARD_BTN_WRITE")
ntpl.tplParseBlock("BOARD_BTN_WRITE")


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("BoardName" , BDV_Name ) _
	,array("actType"   , actType ) _
	,array("pageNo"    , pageNo ) _

	,array("sName"     , sName ) _
	,array("sId"       , sId ) _
	,array("sTitle"    , sTitle ) _
	,array("sContant"  , sContant ) _
	,array("sWord"     , sWord ) _
	,array("PageParams", PageParams ) _


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
	"	,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Form] ) AS [Form] "&_
	"	,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Kind] ) AS [Kind] "&_
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