<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim actType    : actType    = RequestSet("actType","GET","")
Dim Idx        : Idx        = RequestSet("Idx","GET",0)
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
	BoardKey = IIF( BoardKey=0 , BC_FIRST_KEY , BoardKey )
	Call BoardCodeView()

	Call getData()
Call dbclose()

Dim PageParams
PageParams = "pageNo="& pageNo &_
		"&BoardKey="  & BoardKey &_
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
ntpl.setFile "MAIN", "customer/customerV.html"
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

call ntpl.setBlock("MAIN", array("FILE_HIDDEN", "BTN_REPLY"))

If IsNull(FI_File) Or FI_File="" Then
	ntpl.tplBlockDel("FILE_HIDDEN")
Else
	ntpl.tplParseBlock("FILE_HIDDEN")
End If

If BDV_Replyfg = 0 Then 
	ntpl.tplParseBlock("BTN_REPLY")
Else
	ntpl.tplBlockDel("BTN_REPLY")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("BoardName" , BDV_Name ) _
	,array("actType"   , actType ) _
	,array("pageNo"    , pageNo ) _
	,array("BoardKey"  , BoardKey ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sUserId"   , sUserId ) _
	,array("sUserName" , sUserName ) _
	,array("sTitle"    , sTitle ) _
	,array("PageParams", PageParams ) _

	,array("Idx", FI_Idx ) _
	,array("Title", TagDecode(FI_Title) ) _
	,array("Contants", TagDecode(FI_Contants) ) _
	,array("File", FI_File ) _
	,array("Id", FI_Id ) _
	,array("Name", FI_Name ) _
	,array("Secret", FI_Secret ) _
	,array("Pwd", FI_Pwd ) _
	,array("Notice", IIF(FI_Notice="1","checked","") ) _
	,array("Indate", FI_Indate ) _
	,array("Ip", FI_Ip ) _
	,array("downloadUrl", DOWNLOAD_BASE_PATH & FI_File ) _

	,array("UserIdx", FI_UserIdx ) _

	,array("leftMenuOverClass1"   , "" ) _
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
	"/* UPDATE [dbo].[SP_BOARD] SET [RCnt] = [RCnt] + 1 WHERE [Idx] = ?; */ " &_
	"SELECT " &_
	"	 A.[Idx] " &_
	"	,A.[Title] " &_
	"	,A.[Contants] " &_
	"	,A.[File] " &_
	"	,CASE WHEN ISNULL(A.[AdminIdx],0) > 0 THEN C.[Id] ELSE B.[UserId] END AS [Id] " &_
	"	,CASE WHEN ISNULL(A.[AdminIdx],0) > 0 THEN C.[Name] ELSE B.[UserName] END AS [Name] " &_
	"	,A.[Secret] " &_
	"	,A.[Pwd] " &_
	"	,A.[Notice] " &_
	"	,A.[Order] " &_
	"	,A.[Depth] " &_
	"	,A.[Parent] " &_
	"	,A.[Indate] " &_
	"	,A.[Ip] " &_
	"	FROM [dbo].[SP_BOARD] A " &_
	"	LEFT JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx])" &_
	"	LEFT JOIN [dbo].[SP_ADMIN_MEMBER] C ON(A.[AdminIdx] = C.[Idx])" &_
	"WHERE A.[Idx] = ? AND A.[Dellfg] = 0"

	call cmdopen()
	with objCmd
		.CommandText = SQL
		'.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , Idx )
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , Idx )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
%>