<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList   : cntList   = -1
Dim cntTotal  : cntTotal  = 0
Dim rows      : rows      = 20


Dim pageNo     : pageNo    = RequestSet("pageNo","GET",1)
Dim sIndate    : sIndate    = RequestSet("sIndate","GET","")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate","GET","")
Dim sUserId    : sUserId    = RequestSet("sUserId","GET","")
Dim sUserName  : sUserName  = RequestSet("sUserName","GET","")
Dim sHphone3   : sHphone3   = RequestSet("sHphone3","GET","")
Dim sUserBirth : sUserBirth = RequestSet("sUserBirth","GET","")
Dim sState     : sState      = RequestSet("sState","GET","")

Dim sStateOption : sStateOption = ""&_
"<option value="""">선택</option>"&_
"<option value=""0"" "& IIF(sState="0","selected","") &">사용</option>"&_
"<option value=""1"" "& IIF(sState="1","selected","") &">탈퇴</option>"


Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sUserId="    & sUserId &_
		"&sUserName="  & sUserName &_
		"&sHphone3="   & sHphone3 &_
		"&sUserBirth=" & sUserBirth &_
		"&sState="     & sState

Dim pageURL
pageURL	= g_url & "?pageNo=__PAGE__" &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sUserId="    & sUserId &_
		"&sUserName="  & sUserName &_
		"&sHphone3="   & sHphone3 &_
		"&sUserBirth=" & sUserBirth &_
		"&sState="     & sState

checkAdminLogin(g_host & g_url & "?" & PageParams)

Call Expires()
Call dbopen()
	Call GetList()
Call dbclose()

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageURL)

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @pageNo INT , @rows INT;" &_
	"SET @pageNo = ?; " &_
	"SET @rows   = ?; " &_


	"DECLARE @sIndate VARCHAR(10) , @sOutdate VARCHAR(10) , @sUserId VARCHAR(50) , @sUserName VARCHAR(50) , @sHphone3 VARCHAR(4) ,@sUserBirth VARCHAR(8)  ,@sState VARCHAR(8) ;" &_
	"SET @sIndate    = ?; " &_
	"SET @sOutdate   = ?; " &_
	"SET @sUserId    = ?; " &_
	"SET @sUserName  = ?; " &_
	"SET @sHphone3   = ?; " &_
	"SET @sUserBirth = ?; " &_
	"SET @sState      = ?; " &_


	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over ( order by [UserIdx] ) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,[UserIdx] " &_
	"		,[UserId] " &_
	"		,[UserName] " &_
	"		,[UserBirth] " &_
	"		,[UserHphone1] " &_
	"		,[UserHphone2] " &_
	"		,[UserHphone3] " &_
	"		,[UserEmail] " &_
	"		,[UserIndate] " &_
	"		,[UserDelFg] " &_
	"	FROM [dbo].[SP_USER_MEMBER] " &_
	"	WHERE CASE @sIndate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[UserIndate],23) END >= @sIndate " &_
	"	AND   CASE @sOutdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[UserIndate],23) END <= @sOutdate " &_
	"	AND   CASE @sUserId WHEN '' THEN '' ELSE [UserId] END LIKE '%'+@sUserId+'%' " &_
	"	AND   CASE @sUserName WHEN '' THEN '' ELSE [UserName] END LIKE '%'+@sUserName+'%' " &_
	"	AND   CASE @sHphone3 WHEN '' THEN '' ELSE [UserHphone3] END LIKE '%'+@sHphone3+'%' " &_
	"	AND   CASE @sUserBirth WHEN '' THEN '' ELSE [UserBirth] END LIKE '%'+@sUserBirth+'%' " &_
	"	AND   CASE @sState WHEN '' THEN '' ELSE [UserDelFg] END = @sState " &_
	"	/*AND [UserDelFg] = 0*/ " &_
	") " &_
	"SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum desc "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@pageNo"     ,adInteger , adParamInput ,  0 , pageNo )
		.Parameters.Append .CreateParameter( "@rows"       ,adInteger , adParamInput ,  0 , rows )
		.Parameters.Append .CreateParameter( "@sIndate"    ,adVarChar , adParamInput , 10 , sIndate )
		.Parameters.Append .CreateParameter( "@sOutdate"   ,adVarChar , adParamInput , 10 , sOutdate )
		.Parameters.Append .CreateParameter( "@sUserId"    ,adVarChar , adParamInput , 50 , sUserId )
		.Parameters.Append .CreateParameter( "@sUserName"  ,adVarChar , adParamInput , 50 , sUserName )
		.Parameters.Append .CreateParameter( "@sHphone3"   ,adVarChar , adParamInput ,  4 , sHphone3 )
		.Parameters.Append .CreateParameter( "@sUserBirth" ,adVarChar , adParamInput ,  8 , sUserBirth )
		.Parameters.Append .CreateParameter( "@sState"     ,adVarChar , adParamInput ,  8 , sState )
		
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
ntpl.setFile "LEFT", "member/leftMenu.html"
ntpl.setFile "MAIN", "member/memberL.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()

call ntpl.setBlock("MAIN", array("MEMBER_LOOP" , "MEMBER_LOOP_NODATA"))

If cntList > -1 Then 
	'// BLOCK 부분 처리
	for iLoop = 0 to cntList

		ntpl.setBlockReplace array( _
			array("rownum" , arrList(FI_rownum,iLoop)  ), _
			array("UserIdx", arrList(FI_UserIdx,iLoop)), _
			array("UserId", arrList(FI_UserId,iLoop)), _
			array("UserName", arrList(FI_UserName,iLoop) ), _
			array("UserHphone", arrList(FI_UserHphone1,iLoop) &"-"& arrList(FI_UserHphone2,iLoop) &"-"& arrList(FI_UserHphone3,iLoop) ), _
			array("UserEmail", arrList(FI_UserEmail,iLoop) ), _
			array("UserBirth", Mid(arrList(FI_UserBirth,iLoop),1,4) &"-"&Mid(arrList(FI_UserBirth,iLoop),5,2) &"-"&Mid(arrList(FI_UserBirth,iLoop),7,2)  ), _
			array("UserIndate", arrList(FI_UserIndate,iLoop) ), _
			array("UserDelfg", IIF(arrList(FI_UserDelFg,iLoop) = 0 , "사용" , "<span style='color:red;'>탈퇴</span>" ) ) _
		), ""
		ntpl.tplParseBlock("MEMBER_LOOP")
	Next
	ntpl.tplBlockDel("MEMBER_LOOP_NODATA")
Else
	ntpl.tplParseBlock("MEMBER_LOOP_NODATA")
	ntpl.tplBlockDel("MEMBER_LOOP")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("pageList", pagelist ) _
	
	,array("pageNo"    , pageNo ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sUserId"   , sUserId ) _
	,array("sUserName" , sUserName ) _
	,array("sHphone3"  , sHphone3 ) _
	,array("sUserBirth", sUserBirth ) _
	,array("sStateOption" , sStateOption ) _
	,array("sState"    , sState ) _

	,array("PageParams" , PageParams) _

	,array("s1Day"    , Date() ) _
	,array("s7Day"    , Date() -7 ) _
	,array("s30Day"   , Date() -30 ) _
	
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>