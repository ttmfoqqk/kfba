<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1
Dim cntTotal : cntTotal  = 0
Dim rows     : rows      = 20

Dim pageNo     : pageNo    = RequestSet("pageNo" , "GET" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate" , "GET" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate", "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "56")
Dim sState     : sState     = RequestSet("sState"    , "GET" , "")
Dim sId        : sId        = RequestSet("sId"    , "GET" , "")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"    , "GET" , "")
Dim sBirth     : sBirth     = RequestSet("sBirth"    , "GET" , "")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sPcode="     & sPcode &_
		"&sState="     & sState &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sBirth="     & sBirth

Dim pageURL
pageURL	= g_url & "?pageNo=__PAGE__" &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sPcode="     & sPcode &_
		"&sState="     & sState &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sBirth="     & sBirth

checkAdminLogin(g_host & g_url  & "?" & PageParams)

Call Expires()
Call dbopen()

	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )

	Call GetList()
	
	Dim StateOption
	StateOption = "" &_
	"<option value="""">선 택</option>" &_
	"<option value=""1"" "&IIF(sState="1","selected","")&" >접수</option>" &_
	"<option value=""0"" "&IIF(sState="0","selected","")&" >승인</option>" &_
	"<option value=""2"" "&IIF(sState="2","selected","")&" >불합격</option>" 

Call dbclose()


Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageURL)

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @pageNo INT, @rows INT ;" &_
	"SET @pageNo = ?; SET @rows = ?; " &_

	"DECLARE @sIndate VARCHAR(10) , @sOutdate VARCHAR(10) ;" &_
	"DECLARE @sPcode VARCHAR(10) , @sState VARCHAR(3) , @sId VARCHAR(50) , @sName VARCHAR(50) , @sPhone3 VARCHAR(4) , @sBirth VARCHAR(8);" &_
	"SET @sIndate    = ?; " &_
	"SET @sOutdate   = ?; " &_

	"SET @sPcode     = ?; " &_
	"SET @sState     = ?; " &_
	"SET @sId        = ?; " &_
	"SET @sName      = ?; " &_
	"SET @sPhone3    = ?; " &_
	"SET @sBirth     = ?; " &_

	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over (order by A.[Idx] ASC ) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,A.[Idx] " &_
	"		,A.[State] " &_
	"		,CONVERT(VARCHAR,A.[InData],23) AS [InData] " &_
	"		,A.[ProgramKind] " &_
	"		,B.[UserId]" &_
	"		,B.[UserName]" &_
	"		,B.[UserHphone1] " &_
	"		,B.[UserHphone2] " &_
	"		,B.[UserHphone3] " &_
	"		,B.[FirstName] " &_
	"		,B.[LastName] " &_
	"		,C.[Name] AS [ProgramNema]" &_
	"	FROM [dbo].[SP_PROGRAM_JUDGE_APP] A " &_
	"	INNER JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx]) " &_
	"	INNER JOIN [dbo].[SP_COMM_CODE2] C ON(A.[ProgramIdx] = C.[Idx]) " &_
	"	WHERE A.[DellFg] = 0 " &_

	"	AND CASE @sPcode WHEN '' THEN '' ELSE C.[Idx] END = @sPcode "&_  
	"	AND CASE @sState WHEN '' THEN '' ELSE A.[State] END = @sState " &_
	"	AND CASE @sId WHEN '' THEN '' ELSE B.[UserId] END LIKE '%'+@sId+'%' "&_
	"	AND CASE @sName WHEN '' THEN '' ELSE B.[UserName] END LIKE '%'+@sName+'%' " &_
	"	AND CASE @sPhone3 WHEN '' THEN '' ELSE B.[UserHphone3] END LIKE '%'+@sPhone3+'%' "&_
	"	AND CASE @sBirth WHEN '' THEN '' ELSE B.[UserBirth] END LIKE '%'+@sBirth+'%' " &_

	"	AND CASE @sIndate WHEN '' THEN '' ELSE CONVERT(VARCHAR,A.[InData],23) END >= @sIndate " &_
	"	AND CASE @sOutdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,A.[InData],23) END <= @sOutdate " &_

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

		.Parameters.Append .CreateParameter( "@sIndate"    ,adVarChar , adParamInput, 10  , sIndate )
		.Parameters.Append .CreateParameter( "@sOutdate"   ,adVarChar , adParamInput, 10  , sOutdate )
		.Parameters.Append .CreateParameter( "@sPcode"     ,adVarChar , adParamInput, 10  , sPcode )
		.Parameters.Append .CreateParameter( "@sState"     ,adVarChar , adParamInput, 3   , sState )
		.Parameters.Append .CreateParameter( "@sId"        ,adVarChar , adParamInput, 50  , sId )
		.Parameters.Append .CreateParameter( "@sName"      ,adVarChar , adParamInput, 50  , sName )
		.Parameters.Append .CreateParameter( "@sPhone3"    ,adVarChar , adParamInput, 4   , sPhone3 )
		.Parameters.Append .CreateParameter( "@sBirth"     ,adVarChar , adParamInput, 8   , sBirth )
		
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
ntpl.setFile "LEFT", "application/leftMenu.html"
ntpl.setFile "MAIN", "application/judgeL.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()

call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST1","LEFT_MENU_LIST2"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			
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

call ntpl.setBlock("MAIN", array("APPLICATION_LOOP","LOOP_NODATA"))
'// BLOCK 부분 처리

If cntList > -1 Then 
	for iLoop = 0 to cntList

	StateText = ""

	If arrList(FI_State,iLoop) = "0" Then
		StateText = "<font color=""blue"">승인</font>"
	ElseIf arrList(FI_State,iLoop) = "1" Then 
		StateText = "접수"
	ElseIf arrList(FI_State,iLoop) = "2" Then 
		StateText = "<font color=""red"">불합격</font>"
	End If

	PrograName = arrList(FI_ProgramNema,iLoop)

	If arrList(FI_ProgramKind,iLoop) = "1" Then
		PrograName = PrograName & " [필기]"
	ElseIf arrList(FI_ProgramKind,iLoop) = "2" Then
		PrograName = PrograName & " [실기]"
	End If

		ntpl.setBlockReplace array( _
			 array("rownum" , arrList(FI_rownum,iLoop)  ) _
			,array("Idx" , arrList(FI_Idx,iLoop)  ) _
			,array("UserId", arrList(FI_UserId,iLoop) ) _
			,array("UserName", arrList(FI_UserName,iLoop) ) _
			,array("UserPhone", arrList(FI_UserHphone1,iLoop) &"-"& arrList(FI_UserHphone2,iLoop) &"-"& arrList(FI_UserHphone3,iLoop) ) _
			,array("ProgramNema" , PrograName ) _
			,array("InData" , arrList(FI_InData,iLoop)  ) _
			,array("State" , StateText  ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("APPLICATION_LOOP")
	Next
	ntpl.tplBlockDel("LOOP_NODATA")
Else
	ntpl.tplBlockDel("APPLICATION_LOOP")
	ntpl.tplParseBlock("LOOP_NODATA")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("pageList", pagelist ) _
	,array("PageParams", PageParams ) _

	,array("codeOption", codeOption ) _
	,array("StateOption", StateOption ) _
	

	,array("pageNo"   , pageNo ) _
	,array("sIndate"  , sIndate ) _
	,array("sOutdate" , sOutdate ) _
	,array("sPcode"   , sPcode ) _
	,array("sState"   , sState ) _
	,array("sId"      , sId ) _
	,array("sName"    , sName ) _
	,array("sPhone3"  , sPhone3 ) _
	,array("sBirth"   , sBirth ) _

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