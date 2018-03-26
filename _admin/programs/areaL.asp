<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1
Dim cntTotal : cntTotal  = 0
Dim rows     : rows      = 20
Dim pageNo   : pageNo    = RequestSet("pageNo" , "GET" , 1)

Dim sIndate  : sIndate   = RequestSet("sIndate" , "GET" , "")
Dim sOutdate : sOutdate  = RequestSet("sOutdate", "GET" , "")
Dim sName    : sName     = RequestSet("sName"   , "GET" , "")
Dim sAddr    : sAddr     = RequestSet("sAddr"   , "GET" , "")
Dim sCode    : sCode     = RequestSet("sCode"   , "GET" , "")
Dim sTel     : sTel      = RequestSet("sTel"    , "GET" , "")
Dim sPcode   : sPcode    = RequestSet("sPcode"  , "GET" , "56")

Call Expires()
Call dbopen()
	Call common_code_list(17)
	Call GetList()
Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode="     & sPcode &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sName="      & sName &_
		"&sAddr="      & sAddr &_
		"&sCode="      & sCode &_
		"&sTel="       & sTel

Dim pageURL
pageURL	= g_url & "?pageNo=__PAGE__" &_
		"&sPcode="     & sPcode &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sName="      & sName &_
		"&sAddr="      & sAddr &_
		"&sCode="      & sCode &_
		"&sTel="       & sTel

checkAdminLogin(g_host & g_url & "?" & PageParams)

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageURL)

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @pageNo INT, @rows INT ;" &_
	"SET @pageNo = ?; SET @rows = ?; " &_

	"DECLARE @sIndate VARCHAR(10) , @sOutdate VARCHAR(10) , @sName VARCHAR(200) , @sAddr VARCHAR(200) , @sCode VARCHAR(3) , @sTel VARCHAR(50) , @sPcode VARCHAR(10) ;" &_
	"SET @sIndate  = ?; " &_
	"SET @sOutdate = ?; " &_
	"SET @sName    = ?; " &_
	"SET @sAddr    = ?; " &_
	"SET @sCode    = ?; " &_
	"SET @sTel     = ?; " &_
	"SET @sPcode   = ?; " &_

	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over (order by [Idx] ) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,[Idx] " &_
	"		,[Name] " &_
	"		,[Addr] " &_
	"		,[Tel] " &_
	"		,ISNULL([Code],'') AS [Code] " &_
	"		,ISNULL(( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [AddrIdx] ),'') AS [AddrCode]" &_
	"	FROM [dbo].[SP_PROGRAM_AREA] " &_

	"	WHERE [Dellfg] = 0 " &_
	"	AND CASE @sName WHEN '' THEN '' ELSE [Name] END LIKE '%'+@sName+'%' "&_
	"	AND CASE @sAddr WHEN '' THEN '' ELSE [Addr] END LIKE '%'+@sAddr+'%' " &_
	"	AND CASE @sCode WHEN '' THEN '' ELSE [Code] END = @sCode " &_
	"	AND CASE @sTel WHEN '' THEN '' ELSE [Tel] END LIKE '%'+@sTel+'%' " &_
	"	AND CASE @sIndate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[Indate],23) END >= @sIndate " &_
	"	AND CASE @sOutdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[Indate],23) END <= @sOutdate " &_
	"	AND CASE @sPcode WHEN '' THEN '' ELSE [CodeIdx] END = @sPcode " &_

	") " &_
	"SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum DESC "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@pageNo" ,adInteger , adParamInput, 0, pageNo )
		.Parameters.Append .CreateParameter( "@rows"   ,adInteger , adParamInput, 0, rows )

		.Parameters.Append .CreateParameter( "@sIndate"  ,adVarChar , adParamInput, 10  , sIndate )
		.Parameters.Append .CreateParameter( "@sOutdate" ,adVarChar , adParamInput, 10  , sOutdate )
		.Parameters.Append .CreateParameter( "@sName"    ,adVarChar , adParamInput, 200 , sName )
		.Parameters.Append .CreateParameter( "@sAddr"    ,adVarChar , adParamInput, 200 , sAddr )
		.Parameters.Append .CreateParameter( "@sCode"    ,adVarChar , adParamInput, 3   , sCode )
		.Parameters.Append .CreateParameter( "@sTel"     ,adVarChar , adParamInput, 50  , sTel )
		.Parameters.Append .CreateParameter( "@sPcode"   ,adVarChar , adParamInput, 10  , sPcode )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
		cntTotal	= arrList(FI_tcount, 0)
	End If
	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "programs/leftMenu.html"
ntpl.setFile "MAIN", "programs/areaL.html"
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

call ntpl.setBlock("MAIN", array("AREA_LOOP","LOOP_NODATA"))
'// BLOCK 부분 처리

If cntList > -1 Then 
	for iLoop = 0 to cntList

		ntpl.setBlockReplace array( _
			 array("rownum" , arrList(FI_rownum,iLoop)  ) _
			,array("Idx" , arrList(FI_Idx,iLoop)  ) _
			,array("Name", arrList(FI_Name,iLoop) ) _
			,array("Addr", arrList(FI_Addr,iLoop) ) _
			,array("Tel" , arrList(FI_Tel,iLoop)  ) _
			,array("Code" , IIF(arrList(FI_Code,iLoop)="","&nbsp;",arrList(FI_Code,iLoop))  ) _
			,array("AddrCode" , IIF(arrList(FI_AddrCode,iLoop)="","&nbsp;",arrList(FI_AddrCode,iLoop))  ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("AREA_LOOP")
	Next
	ntpl.tplBlockDel("LOOP_NODATA")
Else
	ntpl.tplBlockDel("AREA_LOOP")
	ntpl.tplParseBlock("LOOP_NODATA")
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

	,array("s1Day"    , Date() ) _
	,array("s7Day"    , Date() -7 ) _
	,array("s30Day"   , Date() -30 ) _

	,array("leftMenuOverClass1"   , "" ) _
	,array("leftMenuOverClass2"   , "admin_left_over" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>