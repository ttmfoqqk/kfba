<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList      : cntList      = -1
Dim cntTotal     : cntTotal     = 0
Dim rows         : rows         = 20
Dim pageNo       : pageNo       = RequestSet("pageNo","GET",1)


Dim sIndate      : sIndate      = RequestSet("sIndate"     ,"GET","")
Dim sOutdate     : sOutdate     = RequestSet("sOutdate"    ,"GET","")
Dim sId          : sId          = RequestSet("sId"         ,"GET","")
Dim sName        : sName        = RequestSet("sName"       ,"GET","")

Call Expires()
Call dbopen()
	Call BoardCodeList()
	Call getList()
Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="  & sIndate &_
		"&sOutdate=" & sOutdate &_
		"&sId="      & sId &_
		"&sName="    & sName


Dim pageUrl 
pageUrl = g_url & "?" & "pageNo=__PAGE__" &_
		"&sIndate="  & sIndate &_
		"&sOutdate=" & sOutdate &_
		"&sId="      & sId &_
		"&sName="    & sName

checkAdminLogin(g_host & g_url & "?" & PageParams)

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageUrl)


dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "customer/leftMenu.html"
ntpl.setFile "MAIN", "customer/fJobL.html"
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

call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA"))
If cntList > -1 Then 


	for iLoop = 0 to cntList

		ntpl.setBlockReplace array( _
			array("Number", arrList(FI_rownum,iLoop) ), _
			array("Idx", arrList(FI_Idx,iLoop) ), _
			array("UserName", arrList(FI_UserName,iLoop) ), _
			array("UserId", arrList(FI_UserId,iLoop) ), _
			array("Form", arrList(FI_Form,iLoop) ), _
			array("Kind", arrList(FI_Kind,iLoop) ), _
			array("Pay", arrList(FI_Pay,iLoop) ), _
			array("School", arrList(FI_School,iLoop) ), _
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

	,array("pagelist"  , pagelist ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sId"       , sId ) _
	,array("sName"     , sName ) _
	,array("PageParams", PageParams ) _

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






Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; "& vbCrLf &_

	"DECLARE @pageNo INT , @rows INT;" &_
	"SET @pageNo = ?; " &_
	"SET @rows   = ?; " &_

	"DECLARE @sIndate VARCHAR(10) , @sOutdate VARCHAR(10) , @sId VARCHAR(50) , @sName VARCHAR(50) ;" &_
	"SET @sIndate  = ?; " &_
	"SET @sOutdate = ?; " &_
	"SET @sId      = ?; " &_
	"SET @sName    = ?; " &_
	
	"WITH LIST AS( "&_
	"	SELECT row_number() over (order by A.[Idx] asc) as [rownum]"&_
	"		,count(*) over () as [tcount] "&_
	"		,A.[Idx] "&_
	"		,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Form] ) AS [Form] "&_
	"		,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Kind] ) AS [Kind] "&_
	"		,A.[WorkArea] "&_
	"		,A.[Pay] "&_
	"		,A.[School] "&_
	"		,CONVERT(VARCHAR,A.[InData],23) As [InData] "&_
	"		,B.[UserName]" &_
	"		,B.[UserId]" &_
	"	FROM [dbo].[SP_JOB_USER] A "&_
	"	INNER JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx])"&_

	"	WHERE CASE @sId WHEN '' THEN '' ELSE B.[UserId] END LIKE '%'+@sId+'%' "& vbCrLf &_
	"	AND CASE @sName WHEN '' THEN '' ELSE [UserName] END LIKE '%'+@sName+'%' "& vbCrLf &_
	"	AND CASE @sIndate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[InData],23) END >= @sIndate " &_
	"	AND CASE @sOutdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[InData],23) END <= @sOutdate " &_

	"   AND A.[Dellfg] = 0 "&_
	") SELECT L.* "&_
	"FROM LIST L "&_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) "&_
	"ORDER BY rownum desc; "	

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@pageNo"   ,adInteger , adParamInput ,  0 , pageNo )
		.Parameters.Append .CreateParameter( "@rows"     ,adInteger , adParamInput ,  0 , rows )

		.Parameters.Append .CreateParameter( "@sIndate"  ,adVarChar , adParamInput , 10 , sIndate )
		.Parameters.Append .CreateParameter( "@sOutdate" ,adVarChar , adParamInput , 10 , sOutdate )
		.Parameters.Append .CreateParameter( "@sId"      ,adVarChar , adParamInput , 50 , sId )
		.Parameters.Append .CreateParameter( "@sName"    ,adVarChar , adParamInput , 50 , sName )

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