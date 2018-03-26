<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim arrList
Dim cntList  : cntList  = -1
Dim cntTotal : cntTotal = 0
Dim rows     : rows     = 10
Dim pageNo   : pageNo   = RequestSet("pageNo" ,"GET",1 )
Dim sPcode   : sPcode   = RequestSet("sPcode" ,"GET","")
Dim sACode   : sACode   = RequestSet("sACode" ,"GET","")

Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )

	Call getList()
Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode=" & sPcode &_
		"&sACode=" & sACode

Dim pageUrl 
pageUrl = g_url & "?" & "pageNo=__PAGE__" &_
		"&sPcode=" & sPcode &_
		"&sACode=" & sACode

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageUrl)

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "map/index.html" ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")
call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA"))


If cntList > -1 Then 

	for iLoop = 0 to cntList
		ntpl.setBlockReplace array( _
			 array("Number"     , arrList(FI_rownum,iLoop) ) _
			,array("Idx"        , arrList(FI_Idx,iLoop) ) _
			,array("ProgramName", HtmlTagRemover( arrList(FI_ProgramName,iLoop) , 15) ) _
			,array("Name"       , HtmlTagRemover( arrList(FI_Name,iLoop) , 42) ) _
			,array("Tel"        , HtmlTagRemover( arrList(FI_Tel,iLoop)  , 14) ) _
			,array("Addr"       , HtmlTagRemover( arrList(FI_Addr,iLoop) , 40) ) _
		), ""
		ntpl.tplParseBlock("BOARD_LOOP")
	Next 
	ntpl.tplBlockDel("BOARD_LOOP_NODATA")
Else
	ntpl.tplParseBlock("BOARD_LOOP_NODATA")
	ntpl.tplBlockDel("BOARD_LOOP")
End If

ntpl.tplAssign array(   _
	 array("imgDir"     , TPL_DIR_IMAGES ) _
	,array("codeOption" , codeOption) _
	,array("pagelist"   , pagelist) _
	,array("pageNo"     , pageNo ) _
	,array("sPcode"     , sPcode ) _
	,array("sACode"     , sACode ) _
	,array("PageParams" , PageParams ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing






Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @pageNo INT , @rows INT ,@sPcode VARCHAR(10) , @sACode VARCHAR(10) ;" &_
	"SET @pageNo = ?; " &_
	"SET @rows   = ?; " &_
	"SET @sPcode = ?; " &_
	"SET @sACode = ?; " &_

	"WITH LIST AS( " &_
	"	SELECT row_number() over (order by [Idx]) as [rownum]" &_
	"		,count(*) over () as [tcount] " &_
	"		,[Idx] " &_
	"		,[Name] " &_
	"		,[Addr] " &_
	"		,[Tel] " &_
	"		,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [CodeIdx] ) AS [ProgramName] " &_
	"	FROM [dbo].[SP_PROGRAM_AREA] " &_
	"   WHERE [Dellfg] = 0 " &_
	"   AND ( [Code] IS NOT NULL OR [Code] != '' ) " &_

	"   AND CASE @sPcode WHEN '' THEN '' ELSE [CodeIdx] END = @sPcode " &_
	"   AND CASE @sACode WHEN '' THEN '' ELSE [AddrIdx] END = @sACode " &_
	") SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum desc; "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@pageNo" ,adInteger , adParamInput , 0  , pageNo )
		.Parameters.Append .CreateParameter( "@rows"   ,adInteger , adParamInput , 0  , rows )
		.Parameters.Append .CreateParameter( "@sPcode" ,adVarChar , adParamInput , 20 , sPcode )
		.Parameters.Append .CreateParameter( "@sACode" ,adVarChar , adParamInput , 20 , sACode )
		set objRs = .Execute
	End with
	call cmdclose()

	CALL setFieldIndex(objRs, "FI")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
		cntTotal	= arrList(FI_tcount, 0)
	End If
	Set objRs = Nothing
End Sub
%>