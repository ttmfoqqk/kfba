<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
checkAdminLogin(g_host & g_url)
Dim arrList
Dim cntList   : cntList   = -1
Dim cntTotal  : cntTotal  = 0
Dim rows      : rows      = 20
Dim pageNo    : pageNo    = CInt(IIF(Request.Form("pageNo")="","1",Request.Form("pageNo")))
Dim AdminId   : AdminId   = Request.Form("AdminId")
Dim AdminName : AdminName = Request.Form("AdminName")
Dim Indate    : Indate    = Request.Form("Indate")
Dim Outdate   : Outdate   = Request.Form("Outdate")
Dim pageURL
pageURL	= g_url & "?pageNo=__PAGE__" &_
		"&amp;AdminId="   & AdminId &_
		"&amp;AdminName=" & AdminName &_
		"&amp;Indate="    & Indate &_
		"&amp;Outdate="   & Outdate

Call Expires()
Call dbopen()
	Call GetList()
Call dbclose()

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageURL)

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over (order by [Idx]) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,[Idx] AS [AdminIdx] " &_
	"		,[Id] " &_
	"		,[Pwd] " &_
	"		,[Name] AS [AdminName] " &_
	"		,[pHone1] " &_
	"		,[pHone2] " &_
	"		,[pHone3] " &_
	"		,[Hphone1] " &_
	"		,[Hphone2] " &_
	"		,[Hphone3] " &_
	"		,ISNULL([ExtNum],'') AS [ExtNum] " &_
	"		,ISNULL([DirNum],'') AS [DirNum] " &_
	"		,ISNULL([email],'') AS [email] " &_
	"		,ISNULL([MsgAddr],'') AS [MsgAddr] " &_
	"		,[Bigo] " &_
	"		,CONVERT(VARCHAR,[Indata],23) as [Indata] " &_
	"	FROM [dbo].[SP_ADMIN_MEMBER] " &_
	"	WHERE [Dellfg] = 0 " &_
	") " &_
	"SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(?,?,tcount) AND dbo.fnc_row_to(?,?,tcount) " &_
	"ORDER BY rownum desc "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@pageNo" ,adInteger , adParamInput,  0, pageNo  )
		.Parameters.Append .CreateParameter( "@rows" ,adInteger , adParamInput, 0, rows )
		.Parameters.Append .CreateParameter( "@pageNo" ,adInteger , adParamInput,  0, pageNo  )
		.Parameters.Append .CreateParameter( "@rows" ,adInteger , adParamInput, 0, rows )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
		cntTotal	= arrList(FI_tcount, 0)	' 첫번째에서 행에서 전체 건수 설정.
	End If
	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "company/leftMenu.html"
ntpl.setFile "MAIN", "company/companyMemberL.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

call ntpl.setBlock("MAIN", array("MEMBER_LOOP"))
'// BLOCK 부분 처리
for iLoop = 0 to cntList

	ntpl.setBlockReplace array( _
		array("adminIdx", arrList(FI_AdminIdx,iLoop)), _
		array("adminId", arrList(FI_Id,iLoop)), _
		array("adminName", arrList(FI_AdminName,iLoop) ), _
		array("adminHphone", arrList(FI_Hphone1,iLoop) &"-"& arrList(FI_Hphone2,iLoop) &"-"& arrList(FI_Hphone3,iLoop) ), _
		array("adminEmail", IIF(arrList(FI_Email,iLoop)="","&nbsp;",arrList(FI_Email,iLoop)) ), _
		array("adminMsg", IIF(arrList(FI_MsgAddr,iLoop)="","&nbsp;",arrList(FI_MsgAddr,iLoop)) ), _
		array("adminExtNum", IIF(arrList(FI_ExtNum,iLoop)="","&nbsp;",arrList(FI_ExtNum,iLoop)) ) _
	), ""

	'// MEMBER_LOOP 블럭 누적
	ntpl.tplParseBlock("MEMBER_LOOP")
Next

'//상단메뉴오버
Call topMenuOver()

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("pageList", pagelist ) _
	,array("agree2", FI_Agree2 ) _

	,array("leftMenuOverClass1"   , "" ) _
	,array("leftMenuOverClass2"   , "admin_left_over" ) _
	,array("leftMenuOverClass3"   , "" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>