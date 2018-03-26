<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim arrList , arrNoti
Dim cntList      : cntList      = -1
Dim cntTotal     : cntTotal     = 0
Dim rows         : rows         = 20
Dim pageNo       : pageNo       = RequestSet("pageNo","GET",1)

Dim sName    : sName    = RequestSet("sName"    ,"GET",0 )
Dim sId      : sId      = RequestSet("sId"      ,"GET",0 )
Dim sTitle   : sTitle   = RequestSet("sTitle"   ,"GET",0 )
Dim sContant : sContant = RequestSet("sContant" ,"GET",0 )
Dim sWord    : sWord    = RequestSet("sWord"    ,"GET","")

If sName = 0 And sId = 0 And sTitle = 0 And sContant = 0 Then 
	sName = 1
	sTitle = 1
End If

Call Expires()
Call dbopen()
	Call getList()
Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&BoardKey=" & BoardKey &_
		"&sName="    & sName &_
		"&sId="      & sId &_
		"&sTitle="   & sTitle &_
		"&sContant=" & sContant &_
		"&sWord="    & sWord
Dim pageUrl 
pageUrl = g_url & "?" & "pageNo=__PAGE__" &_
		"&BoardKey=" & BoardKey &_
		"&sName="    & sName &_
		"&sId="      & sId &_
		"&sTitle="   & sTitle &_
		"&sContant=" & sContant &_
		"&sWord="    & sWord


Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageUrl)

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "comunity/fStaffL.html" ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA"))
If cntList > -1 Then 


	for iLoop = 0 to cntList

		ntpl.setBlockReplace array( _
			array("Number", arrList(FI_rownum,iLoop) ), _
			array("Idx", arrList(FI_Idx,iLoop) ), _
			array("Title", HtmlTagRemover(  arrList(FI_Title,iLoop) , 52) ), _
			array("CompanyName", arrList(FI_CompanyName,iLoop) ), _
			array("Form", arrList(FI_Form,iLoop) ), _
			array("Kind", arrList(FI_Kind,iLoop) ), _
			array("Dates", arrList(FI_Dates,iLoop) ), _
			array("InData", arrList(FI_InData,iLoop) ) _
		), ""
		ntpl.tplParseBlock("BOARD_LOOP")
	Next 
	ntpl.tplBlockDel("BOARD_LOOP_NODATA")
Else
	ntpl.tplParseBlock("BOARD_LOOP_NODATA")
	ntpl.tplBlockDel("BOARD_LOOP")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _

	,array("pagelist", pagelist ) _
	,array("sName"    , IIF(sName=1,"checked","") ) _
	,array("sId"      , IIF(sId=1,"checked","") ) _
	,array("sTitle"   , IIF(sTitle=1,"checked","") ) _
	,array("sContant" , IIF(sContant=1,"checked","") ) _
	,array("sWord"    , sWord ) _
	,array("PageParams", PageParams ) _
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
	"DECLARE @pageNo INT , @rows INT;" &_
	"SET @pageNo     = ?; " &_
	"SET @rows      = ?; " &_

	"DECLARE @sId INT , @sName INT , @sTitle INT , @sContant INT , @sWord VARCHAR(MAX) ;" &_
	"SET @sId      = ?; " &_
	"SET @sName    = ?; " &_
	"SET @sTitle   = ?; " &_
	"SET @sContant = ?; " &_
	"SET @sWord    = ?; " &_

	"WITH LIST AS( " &_
	"	SELECT row_number() over (order by [Idx] asc) as [rownum]" &_
	"		,count(*) over () as [tcount] " &_
	"		,[Idx] " &_
	"		,[CompanyName] " &_
	"		,[Title] " &_
	"		,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Form] ) AS [Form] "&_
	"		,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Kind] ) AS [Kind] "&_
	"		,[Dates] " &_
	"		,CONVERT(VARCHAR,[InData],23) AS [InData] " &_
	"	FROM [dbo].[SP_JOB_COMPANY] " &_

	"	WHERE [Dellfg] = 0 " &_
	"   AND ( " &_
	"		CASE @sId WHEN 1 THEN [CompanyName] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"		OR CASE @sName WHEN 1 THEN [OwnerName] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"		OR CASE @sName WHEN 1 THEN [ManagerName] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"		OR CASE @sTitle WHEN 1 THEN [Title] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"		OR CASE @sContant WHEN 1 THEN [Bigo] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"	)" &_
	") SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum desc; "	


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@pageNo"    ,adInteger , adParamInput ,  0   , pageNo )
		.Parameters.Append .CreateParameter( "@rows"      ,adInteger , adParamInput ,  0   , rows )

		.Parameters.Append .CreateParameter( "@sId"       ,adInteger , adParamInput , 0    , sId )
		.Parameters.Append .CreateParameter( "@sName"     ,adInteger , adParamInput , 0    , sName )
		.Parameters.Append .CreateParameter( "@sTitle"    ,adInteger , adParamInput , 0    , sTitle )
		.Parameters.Append .CreateParameter( "@sContant"  ,adInteger , adParamInput , 0    , sContant )
		.Parameters.Append .CreateParameter( "@sWord"     ,adVarChar , adParamInput  , 8000 ,sWord )

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