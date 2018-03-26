<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim Idx      : Idx      = RequestSet("Idx"      ,"GET",0)
Dim BoardKey : BoardKey = RequestSet("BoardKey" ,"GET",0)
Dim pageNo   : pageNo   = RequestSet("pageNo"   ,"GET",1)

Dim sName    : sName    = RequestSet("sName"    ,"GET",0)
Dim sId      : sId      = RequestSet("sId"      ,"GET",0)
Dim sTitle   : sTitle   = RequestSet("sTitle"   ,"GET",0)
Dim sContant : sContant = RequestSet("sContant" ,"GET",0)
Dim sWord    : sWord    = RequestSet("sWord"    ,"GET","")


Call Expires()
Call dbopen()
	BoardKey = IIF( BoardKey=0 , 4 , BoardKey )
	Call BoardCodeView()
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

'읽기권한
If BDV_PmsV = 2 Then 
	Call msgbox("읽기권한이 제한된 게시판 입니다.",true)
ElseIf BDV_PmsV = 1 And (  Isnull( session("UserIdx") ) Or session("UserIdx")=""   ) Then 
	checkLogin( g_host & g_url &"?"&PageParams )
End If


dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "reference/view.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("BOARD_BTN_WRITE"))

'글쓰기 버튼 

If BDV_PmsW = "2" Then 
	ntpl.tplBlockDel("BOARD_BTN_WRITE")
Else
	If FI_UserIdx = Session("UserIdx") And IsNull(FI_AdminIdx) Or FI_AdminIdx="" Or FI_AdminIdx ="0" Then
		ntpl.tplParseBlock("BOARD_BTN_WRITE")
	Else
		ntpl.tplBlockDel("BOARD_BTN_WRITE")
	End If
End If


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("BoardName" , BDV_Name ) _
	,array("actType"   , actType ) _
	,array("pageNo"    , pageNo ) _
	,array("BoardKey"  , BoardKey ) _

	,array("BoardTitle" , BoardTitle ) _
	,array("btnBoardKey1" , btnBoardKey1 ) _
	,array("btnBoardKey2" , btnBoardKey2 ) _

	,array("btnBoardOnOff1" , IIF( int(btnBoardKey1) = int(BoardKey) ,"_on","_off" ) ) _
	,array("btnBoardOnOff2" , IIF( int(btnBoardKey2) = int(BoardKey) ,"_on","_off" ) ) _
	
	,array("sName"     , sName ) _
	,array("sId"       , sId ) _
	,array("sTitle"    , sTitle ) _
	,array("sContant"  , sContant ) _
	,array("sWord"     , sWord ) _
	,array("PageParams", PageParams ) _

	,array("Idx", FI_Idx ) _
	,array("Title", TagDecode(FI_Title) ) _
	,array("Contants", FI_Contants ) _
	,array("File", FI_File ) _
	,array("Id", FI_Id ) _
	,array("Name", FI_Name ) _
	,array("Secret", FI_Secret ) _
	,array("Pwd", FI_Pwd ) _
	,array("Notice", IIF(FI_Notice="1","checked","") ) _
	,array("Indate", FI_Indate ) _
	,array("Ip", FI_Ip ) _
	,array("RCnt", FI_RCnt ) _
	,array("downloadUrl", DOWNLOAD_BASE_PATH & FI_File ) _

	,array("UserIdx", FI_UserIdx ) _
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
	"DECLARE @Idx INT;" &_
	"SET @Idx = ?; " &_
	"UPDATE [dbo].[SP_BOARD] SET [RCnt] = [RCnt] + 1 WHERE [Idx] = @Idx; " &_
	"SELECT " &_
	"	 A.[Idx] " &_
	"	,A.[Title] " &_
	"	,A.[Contants] " &_
	"	,A.[File] " &_
	"	,CASE WHEN A.[UserIdx] IS NULL THEN C.[Id] ELSE B.[UserId] END AS [Id] " &_
	"	,CASE WHEN A.[UserIdx] IS NULL THEN C.[Name] ELSE B.[UserName] END AS [Name] " &_
	"	,A.[Secret] " &_
	"	,A.[Pwd] " &_
	"	,A.[Notice] " &_
	"	,A.[Order] " &_
	"	,A.[Depth] " &_
	"	,A.[Parent] " &_
	"	,convert(varchar(10),A.[Indate],111) as [Indate] " &_
	"	,A.[Ip] " &_
	"	,A.[RCnt] " &_
	"	,A.[UserIdx]" &_
	"	,A.[AdminIdx]" &_
	"	FROM [dbo].[SP_BOARD] A " &_
	"	LEFT JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx])" &_
	"	LEFT JOIN [dbo].[SP_ADMIN_MEMBER] C ON(A.[AdminIdx] = C.[Idx])" &_
	"WHERE A.[Idx] = @Idx AND A.[Dellfg] = 0"

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