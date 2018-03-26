<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1
Dim cntTotal : cntTotal  = 0
Dim rows     : rows      = 20
Dim pageNo   : pageNo    = RequestSet("pageNo" , "GET" , 1)

Dim SHarrList
Dim SHcntList  : SHcntList   = -1

Dim sOnDate    : sOnDate    = RequestSet("sOnDate"  , "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "56")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sKind      : sKind      = RequestSet("sKind"    , "GET" , "")
Dim sClass     : sClass     = RequestSet("sClass"   , "GET" , "")

Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )

	Call GetList()

	Dim SearchOnDateOption : SearchOnDateOption = "<option value="""">선 택</option>"

	for iLoop = 0 to SHcntList
		SearchOnDateOption = SearchOnDateOption & "<option value=""" & SHarrList( SEARCH_OnData,iLoop) & """" & IIF(SHarrList( SEARCH_OnData,iLoop)=sOnDate," selected","") & ">" & SHarrList( SEARCH_OnData,iLoop) & "</option>"
	Next

	Dim KindOption
	KindOption = "" &_
	"<option value="""">선 택</option>" &_
	"<option value=""1"" "&IIF(sKind="1","selected","")&" >필기</option>" &_
	"<option value=""2"" "&IIF(sKind="2","selected","")&" >실기</option>"

	Dim ClassOption
	ClassOption = "" &_
	"<option value="""">선 택</option>" &_
	"<option value=""1"" "&IIF(sClass="1","selected","")&" >1급</option>" &_
	"<option value=""2"" "&IIF(sClass="2","selected","")&" >2급</option>" &_
	"<option value=""3"" "&IIF(sClass="3","selected","")&" >3급</option>"

Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sName="      & sName &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass


Dim pageURL
pageURL	= g_url & "?pageNo=__PAGE__" &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sName="      & sName &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass

checkAdminLogin(g_host & g_url & "?" & PageParams)

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageURL)

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @pageNo INT, @rows INT ;" &_
	"SET @pageNo = ?; SET @rows = ?; " &_

	"DECLARE @sOnDate VARCHAR(10) , @sPcode VARCHAR(10) , @sName VARCHAR(200) , @sKind VARCHAR(5) , @sClass VARCHAR(5) ;" &_
	"SET @sOnDate    = ?; " &_
	"SET @sPcode     = ?; " &_
	"SET @sName      = ?; " &_
	"SET @sKind      = ?; " &_
	"SET @sClass     = ?; " &_


	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over (order by A.[Idx] ) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,A.[Idx] " &_
	"		,A.[OnData] " &_
	"		,ISNULL( A.[Pay],0 ) AS [Pay] " &_
	"		,CONVERT(VARCHAR,A.[InDate],23) AS [InDate] " &_
	"		,A.[Kind] " &_
	"		,A.[Class] " &_
	"		,A.[StartDate] " &_
	"		,A.[EndDate] " &_
	"		,ISNULL(A.[MaxNumber],0) AS [MaxNumber] " &_
	"		,B.[Name] " &_
	"		,ISNULL(C.[Name],'') AS [AreaName]" &_
	"	FROM [dbo].[SP_PROGRAM] A " &_
	"	INNER JOIN [dbo].[SP_COMM_CODE2] B ON(A.[CodeIdx] = B.[Idx]) " &_
	"	LEFT JOIN (" &_
	"		SELECT A.[ProgramIdx],A.[AreaIdx],B.[Name] " &_
	"		FROM [dbo].[SP_PROGRAM_ON_AREA] A " &_
	"		INNER JOIN [dbo].[SP_PROGRAM_AREA] B ON(A.[AreaIdx] = B.[Idx]) " &_
	"		WHERE B.[Dellfg] = 0 " &_
	") C ON(A.[Idx] = C.[ProgramIdx] )" &_
	"	WHERE [Dellfg] = 0 " &_
	"	AND CASE @sPcode WHEN '' THEN '' ELSE B.[Idx] END LIKE '%'+@sPcode+'%' "&_
	"	AND CASE @sName WHEN '' THEN '' ELSE C.[Name] END LIKE '%'+@sName+'%' " &_
	"	AND CASE @sOnDate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[OnData],23) END = @sOnDate " &_
	"	AND CASE @sKind WHEN '' THEN '' ELSE A.[Kind] END = @sKind " &_
	"	AND CASE @sClass WHEN '' THEN '' ELSE A.[Class] END = @sClass " &_
	") " &_
	"SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum DESC; " &_

	"SELECT convert(varchar,A.[OnData],23) AS [OnData] " &_
	"FROM [dbo].[SP_PROGRAM] A " &_
	"INNER JOIN [dbo].[SP_COMM_CODE2] B ON(A.[CodeIdx] = B.[Idx]) " &_
	"where [Dellfg] = 0 " &_
	"AND CASE @sPcode WHEN '' THEN '' ELSE B.[Idx] END = @sPcode " &_
	"group by [OnData] order by [OnData] desc; "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@pageNo"  ,adInteger , adParamInput, 0, pageNo )
		.Parameters.Append .CreateParameter( "@rows"    ,adInteger , adParamInput, 0, rows )
		.Parameters.Append .CreateParameter( "@sOnDate" ,adVarChar , adParamInput, 10  , sOnDate )
		.Parameters.Append .CreateParameter( "@sPcode"  ,adVarChar , adParamInput, 10  , sPcode )
		.Parameters.Append .CreateParameter( "@sName"   ,adVarChar , adParamInput, 200 , sName )
		.Parameters.Append .CreateParameter( "@sKind"   ,adVarChar , adParamInput, 3   , sKind )
		.Parameters.Append .CreateParameter( "@sClass"  ,adVarChar , adParamInput, 3   , sClass )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
		cntTotal	= arrList(FI_tcount, 0)
	End If

	'검정일자 검색용 셀렉트
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "SEARCH")
	If Not(objRs.Eof or objRs.Bof) Then		
		SHarrList = objRs.GetRows()
		SHcntList = UBound(SHarrList, 2)
	End If

	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "programs/leftMenu.html"
ntpl.setFile "MAIN", "programs/programsL.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()

call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST1","LEFT_MENU_LIST2"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("sKey", common_code_arrList(CCODE_sKey,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(common_code_arrList(CCODE_Idx,iLoop))=sPcode,"admin_left_over","" ) ) _
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
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST2")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST2")
End If

call ntpl.setBlock("MAIN", array("PROGRAMS_LOOP","LOOP_NODATA"))
'// BLOCK 부분 처리

If cntList > -1 Then 
	for iLoop = 0 to cntList

		PrograName = arrList(FI_Name,iLoop)

		If arrList(FI_Class,iLoop) = "1" Then
			PrograName = PrograName & " 1급"
		ElseIf arrList(FI_Class,iLoop) = "2" Then
			PrograName = PrograName & " 2급"
		ElseIf arrList(FI_Class,iLoop) = "3" Then
			PrograName = PrograName & " 3급"
		End If

		If arrList(FI_Kind,iLoop) = "1" Then
			PrograName = PrograName & " [필기]"
		ElseIf arrList(FI_Kind,iLoop) = "2" Then
			PrograName = PrograName & " [실기]"		
		End If

		ntpl.setBlockReplace array( _
			 array("rownum" , arrList(FI_rownum,iLoop)  ) _
			,array("Idx" , arrList(FI_Idx,iLoop)  ) _
			,array("Name", PrograName ) _
			,array("OnData", arrList(FI_OnData,iLoop) ) _
			,array("Pay", FormatNumber(arrList(FI_Pay,iLoop),0) & " 원" ) _
			,array("StartDate", arrList(FI_StartDate,iLoop) ) _
			,array("EndDate", arrList(FI_EndDate,iLoop) ) _
			,array("MaxNumber", FormatNumber(arrList(FI_MaxNumber,iLoop),0) & " 명" ) _
			,array("InDate", arrList(FI_InDate,iLoop) ) _
			,array("AreaName", IIF(arrList(FI_AreaName,iLoop)="","&nbsp;",arrList(FI_AreaName,iLoop)) ) _
			
			
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("PROGRAMS_LOOP")
	Next
	ntpl.tplBlockDel("LOOP_NODATA")
Else
	ntpl.tplBlockDel("PROGRAMS_LOOP")
	ntpl.tplParseBlock("LOOP_NODATA")
End If

ntpl.tplAssign array(   _
	 array("imgDir"      , TPL_DIR_IMAGES ) _
	,array("pageList"    , pagelist ) _
	,array("PageParams"  , PageParams ) _
	,array("codeOption"  , codeOption ) _
	,array("pageNo"      , pageNo ) _
	,array("sOnDate"     , sOnDate ) _
	,array("sPcode"      , sPcode ) _
	,array("sName"       , sName ) _
	,array("KindOption"  , KindOption ) _
	,array("ClassOption" , ClassOption ) _

	,array("SearchOnDateOption" , SearchOnDateOption ) _
	,array("leftMenuOverClass1" , "admin_left_over" ) _
	,array("leftMenuOverClass2" , "" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>