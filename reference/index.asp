<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim arrList , arrNoti
Dim cntList  : cntList  = -1
Dim cntNoti  : cntNoti  = -1
Dim cntTotal : cntTotal = 0
Dim rows     : rows     = 10
Dim BoardKey : BoardKey = RequestSet("BoardKey" ,"GET",0)
Dim pageNo   : pageNo   = RequestSet("pageNo"   ,"GET",1 )
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
	'Call BoardCodeList()
	BoardKey = IIF( BoardKey=0 , 4 , BoardKey )
	Call BoardCodeView()
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


If BoardKey = 4 Or BoardKey = 5 Then 
	BoardTitle = 1
	btnBoardKey1 = 4
	btnBoardKey2 = 5
ElseIf BoardKey = 6 Or BoardKey = 7 Then 
	BoardTitle = 2
	btnBoardKey1 = 6
	btnBoardKey2 = 7
ElseIf BoardKey = 8 Or BoardKey = 9 Then 
	BoardTitle = 3
	btnBoardKey1 = 8
	btnBoardKey2 = 9
ElseIf BoardKey = 10 Or BoardKey = 11 Then 
	BoardTitle = 4
	btnBoardKey1 = 10
	btnBoardKey2 = 11
ElseIf BoardKey = 12 Or BoardKey = 13 Then 
	BoardTitle = 5
	btnBoardKey1 = 12
	btnBoardKey2 = 13
End If


'게시판 형식
Dim listFileName
If BDV_Type = "FAQ" Then 
	listFileName = "comunity/faq.html"
Else
	listFileName = "comunity/list.html"
End If


'읽기권한
If BDV_PmsL = 2 Then 
	Call msgbox("읽기권한이 제한된 게시판 입니다.",true)
ElseIf BDV_PmsL = 1 And (  Isnull( session("UserIdx") ) Or session("UserIdx")=""   ) Then 
	checkLogin( g_host & g_url&"?"&PageParams  )
End If

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageUrl)

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "reference/index.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA","BOARD_BTN_WRITE"))

'글쓰기 버튼 

If BDV_PmsW = "2" Then 
	ntpl.tplBlockDel("BOARD_BTN_WRITE")
Else
	ntpl.tplParseBlock("BOARD_BTN_WRITE")
End If


If cntList > -1 Or cntNoti > -1 Then 

	for iLoop = 0 to cntNoti
		ntpl.setBlockReplace array( _
			array("Number", "<img src='{$imgDir}/board/icon_notice.jpg'>" ), _
			array("Idx", arrNoti(NT_Idx,iLoop) ), _
			array("Title", HtmlTagRemover(  arrNoti(NT_Title,iLoop)  , 70 ) ), _
			array("Contants", arrNoti(NT_Contants,iLoop) ), _
			array("Name", arrNoti(NT_Name,iLoop) ), _
			array("Id", arrNoti(NT_Id,iLoop) ), _
			array("Indate", arrNoti(NT_Indate,iLoop) ), _
			array("Rcnt", arrNoti(NT_Rcnt,iLoop) ) _
		), ""
		ntpl.tplParseBlock("BOARD_LOOP")
	Next 

	for iLoop = 0 to cntList
		replyWidth = 0
		replyIcon  = ""
		If arrList(FI_Depth, iLoop) > 0 Then 
			replyWidth = 10 * arrList(FI_Depth, iLoop)
			replyIcon = "<b>→</b> "
		End If

		ntpl.setBlockReplace array( _
			array("Number", arrList(FI_rownum,iLoop) ), _
			array("Idx", arrList(FI_Idx,iLoop) ), _
			array("Title",  replyIcon & HtmlTagRemover(  arrList(FI_Title,iLoop) , 70 ) & IIF(arrList(FI_Secret, iLoop)=1," <img src=""{$imgDir}/board/icon_lock.png"">","") ), _
			array("Contants", arrList(FI_Contants,iLoop) ), _
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
	 array("imgDir"   , TPL_DIR_IMAGES ) _
	,array("BoardName", BDV_Name ) _
	,array("pageNo"   , pageNo ) _
	,array("pagelist" , pagelist ) _
	,array("BoardKey" , BoardKey ) _
	,array("BoardTitle" , BoardTitle ) _
	,array("btnBoardKey1" , btnBoardKey1 ) _
	,array("btnBoardKey2" , btnBoardKey2 ) _

	,array("btnBoardOnOff1" , IIF( int(btnBoardKey1) = int(BoardKey) ,"_on","_off" ) ) _
	,array("btnBoardOnOff2" , IIF( int(btnBoardKey2) = int(BoardKey) ,"_on","_off" ) ) _


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
	"DECLARE @pageNo INT , @rows INT ,@BoardKey INT ;" &_
	"SET @pageNo     = ?; " &_
	"SET @rows      = ?; " &_
	"SET @BoardKey  = ?; " &_

	"DECLARE @sId INT , @sName INT , @sTitle INT , @sContant INT , @sWord VARCHAR(MAX) ;" &_
	"SET @sId      = ?; " &_
	"SET @sName    = ?; " &_
	"SET @sTitle   = ?; " &_
	"SET @sContant = ?; " &_
	"SET @sWord    = ?; " &_

	"SELECT " &_
	"	 A.[Idx] " &_
	"	,A.[Title] " &_
	"	,A.[Contants] " &_
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

	"WITH LIST AS( " & vbCrLf &_
	"	SELECT row_number() over (order by A.[Parent] asc, A.[Order] desc) as [rownum]" &_
	"		,count(*) over () as [tcount] " &_
	"		,A.[Idx] " &_
	"		,A.[Title] " &_
	"		,A.[Contants] " &_
	"		,CASE WHEN ISNULL(A.[UserIdx],0) > 0 THEN B.[UserId] ELSE C.[Id] END AS [Id] " &_
	"		,CASE WHEN ISNULL(A.[UserIdx],0) > 0 THEN B.[UserName] ELSE C.[Name] END AS [Name] " &_
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
	"   WHERE A.[Dellfg] = 0 " &_
	"   AND A.[BoardKey] = @BoardKey " &_
	"   AND ( " &_
	"		CASE @sId WHEN 1 THEN [Id] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"		OR CASE @sName WHEN 1 THEN [Name] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"		OR CASE @sTitle WHEN 1 THEN [Title] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"		OR CASE @sContant WHEN 1 THEN [Contants] ELSE '' END LIKE '%'+@sWord+'%' " &_
	"	)" &_
	") SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum desc; "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@pageNo"    ,adInteger , adParamInput ,  0  , pageNo )
		.Parameters.Append .CreateParameter( "@rows"      ,adInteger , adParamInput ,  0  , rows )
		.Parameters.Append .CreateParameter( "@BoardKey"  ,adInteger , adParamInput ,  0  , BoardKey )

		.Parameters.Append .CreateParameter( "@sId"       ,adInteger , adParamInput , 0   , sId )
		.Parameters.Append .CreateParameter( "@sName"     ,adInteger , adParamInput , 0   , sName )
		.Parameters.Append .CreateParameter( "@sTitle"    ,adInteger , adParamInput , 0   , sTitle )
		.Parameters.Append .CreateParameter( "@sContant"  ,adInteger , adParamInput , 0   , sContant )
		.Parameters.Append .CreateParameter( "@sWord"    ,adVarChar , adParamInput  , 8000,sWord )

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