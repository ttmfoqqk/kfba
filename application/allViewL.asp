<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim arrList
Dim cntList   : cntList   = -1
Dim cntTotal  : cntTotal  = 0
Dim rows      : rows      = 10

Dim SHarrList
Dim SHcntList : SHcntList = -1

Dim pageNo    : pageNo    = RequestSet("pageNo"  ,"GET" ,1 )
Dim sOnDate   : sOnDate   = RequestSet("sOnDate" ,"GET" ,"")
Dim sPcode    : sPcode    = RequestSet("sPcode"  ,"GET" ,"")
Dim sKind     : sKind     = RequestSet("sKind"   ,"GET" ,"")
Dim sClass    : sClass    = RequestSet("sClass"  ,"GET" ,"")

Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )
	Call getList()

	Dim SearchOnDateOption : SearchOnDateOption = "<option value="""">선 택</option>"

	for iLoop = 0 to SHcntList
		SearchOnDateOption = SearchOnDateOption & "<option value=""" & SHarrList( SEARCH_OnData,iLoop) & """" & IIF(SHarrList( SEARCH_OnData,iLoop)=sOnDate," selected","") & ">" & SHarrList( SEARCH_OnData,iLoop) & "</option>"
	Next

Call dbclose()

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

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sOnDate=" & sOnDate &_
		"&sPcode="  & sPcode &_
		"&sKind="   & sKind &_
		"&sClass="  & sClass

Dim pageUrl 
pageUrl = g_url & "?" & "pageNo=__PAGE__" &_
		"&sOnDate=" & sOnDate &_
		"&sPcode="  & sPcode &_
		"&sKind="   & sKind &_
		"&sClass="  & sClass

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageUrl)

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "application/allViewL.html" ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")
call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA"))


If cntList > -1 Then 

	for iLoop = 0 to cntList

		PrograName = arrList(FI_ProgramName,iLoop)

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

		ProgramLink = "write.asp?applicationKey=" & arrList(FI_CodeIdx,iLoop)
		' 마감
		If arrList(FI_EndDate,iLoop) < Left(Now(),10) Then
			ProgramLink = "javascript:void(alert('응시 마감되었습니다.'))"
		End If
		' 접수전
		If arrList(FI_StartDate,iLoop) > Left(Now(),10) Then 
			ProgramLink = "javascript:void(alert('응시 접수기간이 아닙니다.'))"
		End If
		' 인원제한
		If arrList(FI_MaxNumber,iLoop) <= arrList(FI_CNT_APP,iLoop) Then 
			ProgramLink = "javascript:void(alert('응시 정원초과!'))"
		End If

		ntpl.setBlockReplace array( _
			 array("Number"     , arrList(FI_rownum,iLoop) ) _
			,array("Idx"        , arrList(FI_Idx,iLoop) ) _
			,array("Link"       , ProgramLink ) _
			,array("ProgramName", PrograName ) _
			,array("OnData"     , arrList(FI_OnData,iLoop) ) _
			,array("StartDate"  , arrList(FI_StartDate,iLoop) ) _
			,array("EndDate"    , arrList(FI_EndDate,iLoop) ) _
			,array("Pay"        , FormatNumber(arrList(FI_Pay,iLoop),0) ) _
			,array("MaxNumber"  , arrList(FI_MaxNumber,iLoop) ) _
			,array("AreaName"   , HtmlTagRemover( arrList(FI_AreaName,iLoop) , 25 ) ) _
			
		), ""
		ntpl.tplParseBlock("BOARD_LOOP")
	Next 
	ntpl.tplBlockDel("BOARD_LOOP_NODATA")
Else
	ntpl.tplParseBlock("BOARD_LOOP_NODATA")
	ntpl.tplBlockDel("BOARD_LOOP")
End If

ntpl.tplAssign array( _
	 array("imgDir"     , TPL_DIR_IMAGES ) _
	,array("codeOption" , codeOption) _
	,array("KindOption" , KindOption ) _
	,array("ClassOption", ClassOption ) _
	,array("SearchOnDateOption", SearchOnDateOption ) _

	,array("pagelist"   , pagelist) _
	,array("pageNo"     , pageNo ) _
	,array("sPcode"     , sPcode ) _
	,array("sACode"     , sACode ) _
	,array("PageParams" , PageParams ) _
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
	"DECLARE @pageNo INT , @rows INT ,@sPcode VARCHAR(10),@sKind VARCHAR(2),@sClass VARCHAR(2),@sOnDate VARCHAR(7) ;" &_
	"SET @pageNo  = ?; " &_
	"SET @rows    = ?; " &_
	"SET @sPcode  = ?; " &_
	"SET @sKind   = ?; " &_
	"SET @sClass  = ?; " &_
	"SET @sOnDate = ?; " &_

	"WITH LIST AS( " &_
	"	SELECT row_number() over (order by [OnData] , [Idx] desc ) as [rownum]" &_
	"		,count(*) over () as [tcount] " &_
	"		,[Idx] " &_
	"		,[CodeIdx] " &_
	"		,[OnData] " &_
	"		,ISNULL( [Pay] , 0 ) AS [Pay] " &_
	"		,CONVERT(varchar(10),[StartDate],23) AS [StartDate] " &_
	"		,CONVERT(varchar(10),[EndDate],23) AS [EndDate] " &_
	"		,ISNULL( [MaxNumber] , 0 ) AS [MaxNumber] " &_
	"		,[Kind] " &_
	"		,[Class] " &_
	"		,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = A.[CodeIdx] ) AS [ProgramName] " &_
	"		,ISNULL(B.[CNT_APP],0) AS [CNT_APP] " &_
	"		,C.[Name] AS [AreaName] " &_
	"	FROM [dbo].[SP_PROGRAM] A " &_

	"	INNER JOIN ( " &_
	"		SELECT " &_
	"			 A.[ProgramIdx] " &_
	"			,B.[Name] " &_
	"		FROM [dbo].[SP_PROGRAM_ON_AREA] A " &_
	"		INNER JOIN [dbo].[SP_PROGRAM_AREA] B ON(A.[AreaIdx]=B.[Idx]) " &_
	"	) C ON(A.[Idx] = C.[ProgramIdx] ) " &_

	"	LEFT JOIN ( " &_
	"		SELECT " &_
	"			 [ProgramIdx] " &_
	"			,COUNT(*) AS [CNT_APP] " &_
	"		FROM [dbo].[SP_PROGRAM_APP] " &_
	"		WHERE [State] != 2 " &_
	"		GROUP BY [ProgramIdx] " &_
	"	) B ON(A.[Idx] = B.[ProgramIdx] ) " &_

	"   WHERE [Dellfg] = 0 " &_
	"   AND CASE @sPcode WHEN '' THEN '' ELSE [CodeIdx] END = @sPcode " &_
	"   AND CASE @sKind WHEN '' THEN '' ELSE [Kind] END = @sKind " &_
	"   AND CASE @sClass WHEN '' THEN '' ELSE [Class] END = @sClass " &_
	"   AND CASE @sOnDate WHEN '' THEN '' ELSE convert(varchar(7),A.[OnData],23) END = @sOnDate " &_
	") SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum desc; " &_

	"SELECT convert(varchar(7),A.[OnData],23) AS [OnData] " &_
	"FROM [dbo].[SP_PROGRAM] A " &_
	"INNER JOIN [dbo].[SP_COMM_CODE2] B ON(A.[CodeIdx] = B.[Idx]) " &_
	"where [Dellfg] = 0 " &_
	"/*AND CASE @sPcode WHEN '' THEN '' ELSE B.[Idx] END = @sPcode*/ " &_
	"group by convert(varchar(7),A.[OnData],23) order by [OnData] desc; "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@pageNo"  ,adInteger , adParamInput , 0  , pageNo )
		.Parameters.Append .CreateParameter( "@rows"    ,adInteger , adParamInput , 0  , rows )
		.Parameters.Append .CreateParameter( "@sPcode"  ,adVarChar , adParamInput , 20 , sPcode )
		.Parameters.Append .CreateParameter( "@sKind"   ,adVarChar , adParamInput , 2  , sKind )
		.Parameters.Append .CreateParameter( "@sClass"  ,adVarChar , adParamInput , 2  , sClass )
		.Parameters.Append .CreateParameter( "@sOnDate" ,adVarChar , adParamInput , 7  , sOnDate )
		set objRs = .Execute
	End with
	call cmdclose()

	CALL setFieldIndex(objRs, "FI")
	If Not(objRs.Eof or objRs.Bof) Then		
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

	Set objRs = Nothing
End Sub
%>