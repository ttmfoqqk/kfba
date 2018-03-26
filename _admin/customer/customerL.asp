<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%


Dim arrList , arrNoti
Dim cntList    : cntList  = -1
Dim cntNoti    : cntNoti  = -1
Dim cntTotal   : cntTotal = 0
Dim rows       : rows     = 20

Dim pageNo     : pageNo     = RequestSet("pageNo","GET",1)
Dim BoardKey   : BoardKey   = RequestSet("BoardKey","GET",0)
Dim sIndate    : sIndate    = RequestSet("sIndate","GET","")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate","GET","")
Dim sUserId    : sUserId    = RequestSet("sUserId","GET","")
Dim sUserName  : sUserName  = RequestSet("sUserName","GET","")
Dim sTitle     : sTitle     = RequestSet("sTitle","GET","")


Call Expires()
Call dbopen()
	Call BoardCodeList()
	BoardKey = IIF( BoardKey=0 , CStr(BC_FIRST_KEY) , BoardKey )
	Call BoardCodeView()
	Call getList()
Call dbclose()

Dim PageParams
PageParams = "pageNo="& pageNo &_
		"&BoardKey="  & BoardKey &_
		"&sIndate="   & sIndate &_
		"&sOutdate="  & sOutdate &_
		"&sUserId="   & sUserId &_
		"&sUserName=" & sUserName &_
		"&sTitle="    & sTitle
Dim pageUrl 
pageUrl = g_url & "?" & "pageNo=__PAGE__" &_
		"&BoardKey="  & BoardKey &_
		"&sIndate="   & sIndate &_
		"&sOutdate="  & sOutdate &_
		"&sUserId="   & sUserId &_
		"&sUserName=" & sUserName &_
		"&sTitle="    & sTitle

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageUrl)

checkAdminLogin(g_host & g_url & "?" & PageParams)



dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "customer/leftMenu.html"
ntpl.setFile "MAIN", "customer/customerL.html"
ntpl.setFile "FOOTER", "_inc/footer.html"
'//상단메뉴오버
Call topMenuOver()
'//왼쪽메뉴 설정
call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST"))
If BC_CNT_LIST > -1 Then 
	for iLoop = 0 to BC_CNT_LIST
		ntpl.setBlockReplace array( _
			 array("Idx", BC_ARRY_LIST(BDL_Idx,iLoop) ) _
			,array("Name", BC_ARRY_LIST(BDL_Name,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(BC_ARRY_LIST(BDL_Idx,iLoop))=BoardKey,"admin_left_over","" ) ) _
		), ""
		ntpl.tplParseBlock("LEFT_MENU_LIST")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST")
End If
'//왼쪽메뉴 설정 끝

call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA"))
If cntList > -1 Or cntNoti > -1 Then 

	for iLoop = 0 to cntNoti
		ntpl.setBlockReplace array( _
			array("Number", "<span style='color:red'>공지</span>" ), _
			array("Idx", arrNoti(NT_Idx,iLoop) ), _
			array("Title", arrNoti(NT_Title,iLoop) ), _
			array("Name", arrNoti(NT_Name,iLoop) ), _
			array("Id", arrNoti(NT_Id,iLoop) ), _
			array("Indate", arrNoti(NT_Indate,iLoop) ), _
			array("Rcnt", arrNoti(NT_Rcnt,iLoop) ) _
		), ""
		ntpl.tplParseBlock("BOARD_LOOP")
	Next 

	for iLoop = 0 to cntList
		replyWidth = 0
		If arrList(FI_Depth, iLoop) > 0 Then 
			replyWidth = 10 * arrList(FI_Depth, iLoop)
			arrList(FI_Title,iLoop) = "└ " & arrList(FI_Title,iLoop)
		End If

		ntpl.setBlockReplace array( _
			array("Number", arrList(FI_rownum,iLoop) ), _
			array("Idx", arrList(FI_Idx,iLoop) ), _
			array("Title", arrList(FI_Title,iLoop) ), _
			array("Name", arrList(FI_Name,iLoop) ), _
			array("Id", arrList(FI_Id,iLoop) ), _
			array("Indate", arrList(FI_Indate,iLoop) ), _
			array("Rcnt", arrList(FI_Rcnt,iLoop) ), _
			array("replyWidth", replyWidth ) _
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
	,array("BoardName", BDV_Name ) _
	,array("pagelist", pagelist ) _
	,array("BoardKey", BoardKey ) _
	
	,array("pageNo"    , pageNo ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sUserId"   , sUserId ) _
	,array("sUserName" , sUserName ) _
	,array("sTitle"    , sTitle ) _

	,array("PageParams" , PageParams) _

	,array("s1Day"    , Date() ) _
	,array("s7Day"    , Date() -7 ) _
	,array("s30Day"   , Date() -30 ) _
	,array("leftMenuOverClass1"   , "" ) _
	,array("leftMenuOverClass2"   , "" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing






Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_

	"DECLARE @pageNo INT , @rows INT;" &_
	"SET @pageNo = ?; " &_
	"SET @rows   = ?; " &_

	"DECLARE @BoardKey INT , @Id VARCHAR(50) , @Name VARCHAR(50) , @Title VARCHAR(200) , @Indate VARCHAR(10) , @Outdate VARCHAR(10) ;" &_
	"SET @BoardKey = ?; " &_
	"SET @Id       = ?; " &_
	"SET @Name     = ?; " &_
	"SET @Title    = ?; " &_
	"SET @Indate   = ?; " &_
	"SET @Outdate = ?; " &_

	"SELECT " &_
	"	 A.[Idx] " &_
	"	,A.[Title] " &_
	"	,B.[Id] " &_
	"	,B.[Name] " &_
	"	,A.[RCnt] " &_
	"	,A.[CmCnt] " &_
	"	,CONVERT(VARCHAR,A.[Indate],23) AS [Indate] " &_
	"FROM [dbo].[SP_BOARD] A " &_
	"INNER JOIN [dbo].[SP_ADMIN_MEMBER] B ON(A.[AdminIdx] = B.[Idx])" &_
	"WHERE [Notice] = 1 AND A.[Dellfg] = 0" &_
	"AND A.[BoardKey] = @BoardKey " &_
	"ORDER BY A.[Idx] DESC; " &_

	"WITH LIST AS( "  &_
	"	SELECT row_number() over (order by A.[Parent] asc, A.[Order] desc) as [rownum]" &_
	"		,count(*) over () as [tcount] " &_
	"		,A.[Idx] " &_
	"		,A.[Title] " &_
	"		,CASE WHEN ISNULL(A.[AdminIdx],0) > 0 THEN C.[Id] ELSE B.[UserId] END AS [Id] " &_
	"		,CASE WHEN ISNULL(A.[AdminIdx],0) > 0 THEN C.[Name] ELSE B.[UserName] END AS [Name] " &_
	"		,A.[Secret] " &_
	"		,A.[Order] " &_
	"		,A.[Depth] " &_
	"		,A.[Parent] " &_
	"		,A.[RCnt] " &_
	"		,A.[CmCnt] " &_
	"		,CONVERT(VARCHAR,A.[Indate],23) AS [Indate] " &_
	"	FROM [dbo].[SP_BOARD] A " &_
	"	LEFT JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx])" &_
	"	LEFT JOIN [dbo].[SP_ADMIN_MEMBER] C ON(A.[AdminIdx] = C.[Idx])" &_
	"	WHERE CASE @Id WHEN '' THEN '' ELSE [Id] END LIKE '%'+@Id+'%' " &_
	"	AND CASE @Name WHEN '' THEN '' ELSE [Name] END LIKE '%'+@Name+'%' " &_
	"	AND CASE @Title WHEN '' THEN '' ELSE [Title] END LIKE '%'+@Title+'%' " &_
	"	AND CASE @Indate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[Indate],23) END >= @Indate " &_
	"	AND CASE @Outdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[Indate],23) END <= @Outdate " &_
	"   AND A.[Dellfg] = 0 " &_
	"   AND A.[BoardKey] = @BoardKey " &_
	") SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum desc; "	

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@pageNo"   ,adInteger , adParamInput ,  0 , pageNo )
		.Parameters.Append .CreateParameter( "@rows"     ,adInteger , adParamInput ,  0 , rows )

		.Parameters.Append .CreateParameter( "@BoardKey" ,adInteger , adParamInput ,  0 , BoardKey )
		.Parameters.Append .CreateParameter( "@Id"       ,adVarChar , adParamInput , 50 , sUserId )
		.Parameters.Append .CreateParameter( "@Name"     ,adVarChar , adParamInput , 50 , sUserName )
		.Parameters.Append .CreateParameter( "@Title"    ,adVarChar , adParamInput , 200 ,sTitle )
		.Parameters.Append .CreateParameter( "@Indate"   ,adVarChar , adParamInput , 10 , sIndate )
		.Parameters.Append .CreateParameter( "@Outdate"  ,adVarChar , adParamInput , 10 , sOutdate )
		set objRs = .Execute
	End with
	call cmdclose()
	'공지사항 리스트
	CALL setFieldIndex(objRs, "NT")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrNoti		= objRs.GetRows()
		cntNoti		= UBound(arrNoti, 2)
	End If
	'게시글
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "FI")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
		cntTotal	= arrList(FI_tcount, 0)
	End If
	Set objRs = Nothing
End Sub

%>