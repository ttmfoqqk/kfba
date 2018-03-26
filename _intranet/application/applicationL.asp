<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1
Dim cntTotal : cntTotal  = 0
Dim rows     : rows      = 20

Dim SHarrList
Dim SHcntList  : SHcntList   = -1

Dim pageNo     : pageNo     = RequestSet("pageNo"   , "GET" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate"  , "GET" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate" , "GET" , "")
Dim sOnDate    : sOnDate    = RequestSet("sOnDate"  , "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "56")
Dim sArea      : sArea      = RequestSet("sArea"    , "GET" , "")

Dim sId        : sId        = RequestSet("sId"      , "GET" , "")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"  , "GET" , "")
Dim sState     : sState     = RequestSet("sState"   , "GET" , "")
Dim sSnumber   : sSnumber   = RequestSet("sSnumber" , "GET" , "")
Dim sKind      : sKind      = RequestSet("sKind"    , "GET" , "")
Dim sClass     : sClass     = RequestSet("sClass"   , "GET" , "")

Dim sOnTime    : sOnTime    = RequestSet("sOnTime"  , "GET" , "")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sArea="      & sArea &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sState="     & sState &_
		"&sSnumber="   & sSnumber &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass &_
		"&sOnTime="    & sOnTime

Dim pageURL
pageURL	= g_url & "?pageNo=__PAGE__" &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sArea="      & sArea &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sState="     & sState &_
		"&sSnumber="   & sSnumber &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass &_
		"&sOnTime="    & sOnTime


checkAdminLogin(g_host & g_url & "?" & PageParams)

Dim hoursOption : hoursOption = "<option value="""">선택</option>"
for iLoop = 7 to 20
	tmp_value = IIF( iLoop < 10 , "0" & iLoop , iLoop )
	tmp_tt    = IIF( iLoop < 12 , "오전", "오후" )
	tmp_hh    = IIF( iLoop < 13 , iLoop , iLoop - 12 )
	'tmp_hh    = iLoop
	hoursOption = hoursOption & "<option value=""" & tmp_value & """" & IIF(sOnTime=Trim(tmp_value)," selected","") & ">" & tmp_tt & " " & tmp_hh & "</option>"
Next 


Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )

	Dim StateOption
	StateOption = "" &_
	"<option value="""">선 택</option>" &_
	"<option value=""1"" "&IIF(sState="1","selected","")&" >입금대기</option>" &_
	"<option value=""0"" "&IIF(sState="0","selected","")&" >접수완료</option>" &_
	"<option value=""2"" "&IIF(sState="2","selected","")&" >접수취소</option>" &_
	"<option value=""3"" "&IIF(sState="3","selected","")&" >불합격</option>" &_
	"<option value=""4"" "&IIF(sState="4","selected","")&" >미응시(불합격)</option>" &_
	"<option value=""10"" "&IIF(sState="10","selected","")&" >합격</option>"

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
	"<option value=""3"" "&IIF(sClass="3","selected","")&" >SPECIAL</option>"


	Call GetList()

	Dim SearchOnDateOption : SearchOnDateOption = "<option value="""">선 택</option>"

	for iLoop = 0 to SHcntList
		SearchOnDateOption = SearchOnDateOption & "<option value=""" & SHarrList( SEARCH_OnData,iLoop) & """" & IIF(SHarrList( SEARCH_OnData,iLoop)=sOnDate," selected","") & ">" & SHarrList( SEARCH_OnData,iLoop) & "</option>"
	Next

Call dbclose()



Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageURL)

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @pageNo INT, @rows INT ;" &_
	"SET @pageNo = ?; SET @rows = ?; " &_

	"DECLARE @sIndate VARCHAR(10) , @sOutdate VARCHAR(10) , @sOnDate VARCHAR(10);" &_
	"DECLARE @sPcode VARCHAR(10) , @sArea VARCHAR(200) , @sId VARCHAR(50) , @sName VARCHAR(50) , @sPhone3 VARCHAR(4) , @sState VARCHAR(3) , @sSnumber VARCHAR(13),@Judge_Idx INT ,@sKind VARCHAR(5) , @sClass VARCHAR(5),@sOnTime VARCHAR(2) ;" &_
	"SET @sIndate    = ?; " &_
	"SET @sOutdate   = ?; " &_
	"SET @sOnDate    = ?; " &_

	"SET @sPcode     = ?; " &_
	"SET @sArea      = ?; " &_
	"SET @sId        = ?; " &_
	"SET @sName      = ?; " &_
	"SET @sPhone3    = ?; " &_
	"SET @sState     = ?; " &_
	"SET @sSnumber   = ?; " &_
	"SET @Judge_Idx  = ?;" &_
	"SET @sKind      = ?; " &_
	"SET @sClass     = ?; " &_
	"SET @sOnTime    = ?; " &_

	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over (order by A.[Idx] ASC ) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,A.[Idx] " &_
	"		,A.[State] " &_
	"		,A.[Snumber] " &_
	"		,CONVERT(VARCHAR,A.[InData],23) AS [InData] " &_
	"		,B.[UserId]" &_
	"		,B.[UserName]" &_
	"		,B.[UserHphone1] " &_
	"		,B.[UserHphone2] " &_
	"		,B.[UserHphone3] " &_
	"		,B.[FirstName] " &_
	"		,B.[LastName] " &_
	"		,C.[Name] AS [ProgramNema]" &_
	"		,C.[Kind]" &_
	"		,C.[Class]" &_
	"		,convert( varchar, C.[OnData],23 ) AS [OnData] " &_
	"		,D.[Name] AS [AreaName] " &_
	"	FROM [dbo].[SP_PROGRAM_APP] A " &_
	"	INNER JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx]) " &_
	"	INNER JOIN ( " &_
	"		SELECT A.[Idx],A.[OnData],A.[Kind],A.[Class],B.[Idx] AS [ProgramIdx]  ,B.[Name] FROM [dbo].[SP_PROGRAM] A INNER JOIN [dbo].[SP_COMM_CODE2] B ON(A.[CodeIdx] = B.[Idx])" &_
	"	) C ON(A.[ProgramIdx] = C.[Idx]) " &_

	"	INNER JOIN [dbo].[SP_PROGRAM_AREA] D ON(A.[AreaIdx] = D.[Idx] ) " &_

	"	WHERE CASE @sPcode WHEN '' THEN '' ELSE C.[ProgramIdx] END = @sPcode "&_
	"	AND CASE @sArea WHEN '' THEN '' ELSE D.[Name] END LIKE '%'+@sArea+'%' " &_
	"	AND CASE @sId WHEN '' THEN '' ELSE B.[UserId] END LIKE '%'+@sId+'%' "&_
	"	AND CASE @sName WHEN '' THEN '' ELSE B.[UserName] END LIKE '%'+@sName+'%' " &_
	"	AND CASE @sPhone3 WHEN '' THEN '' ELSE B.[UserHphone3] END LIKE '%'+@sPhone3+'%' "&_
	"	AND CASE @sState WHEN '' THEN '' ELSE A.[State] END = @sState " &_
	"	AND CASE @sSnumber WHEN '' THEN '' ELSE A.[Snumber] END = @sSnumber " &_
	"	AND CASE @sKind WHEN '' THEN '' ELSE C.[Kind] END = @sKind " &_
	"	AND CASE @sClass WHEN '' THEN '' ELSE C.[Class] END = @sClass " &_

	"	AND CASE @sIndate WHEN '' THEN '' ELSE CONVERT(VARCHAR,A.[InData],23) END >= @sIndate " &_
	"	AND CASE @sOutdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,A.[InData],23) END <= @sOutdate " &_
	"	AND CASE @sOnDate WHEN '' THEN '' ELSE CONVERT(VARCHAR,C.[OnData],23) END = @sOnDate " &_
	"	AND CASE @sOnTime WHEN '' THEN '' ELSE CONVERT(VARCHAR(2),C.[OnData],108) END = @sOnTime " &_

	"	AND A.[AreaIdx] = @Judge_Idx " &_

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
	"group by convert(varchar,A.[OnData],23) order by [OnData] desc; "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@pageNo" ,adInteger , adParamInput, 0, pageNo )
		.Parameters.Append .CreateParameter( "@rows"   ,adInteger , adParamInput, 0, rows )

		.Parameters.Append .CreateParameter( "@sIndate"    ,adVarChar , adParamInput, 10  , sIndate )
		.Parameters.Append .CreateParameter( "@sOutdate"   ,adVarChar , adParamInput, 10  , sOutdate )
		.Parameters.Append .CreateParameter( "@sOnDate"    ,adVarChar , adParamInput, 10  , sOnDate )

		.Parameters.Append .CreateParameter( "@sPcode"     ,adVarChar , adParamInput, 10  , sPcode )
		.Parameters.Append .CreateParameter( "@sArea"      ,adVarChar , adParamInput, 200 , sArea )
		.Parameters.Append .CreateParameter( "@sId"        ,adVarChar , adParamInput, 50  , sId )
		.Parameters.Append .CreateParameter( "@sName"      ,adVarChar , adParamInput, 50  , sName )
		.Parameters.Append .CreateParameter( "@sPhone3"    ,adVarChar , adParamInput, 4   , sPhone3 )
		.Parameters.Append .CreateParameter( "@sState"     ,adVarChar , adParamInput, 3   , sState )
		.Parameters.Append .CreateParameter( "@sSnumber"   ,adVarChar , adParamInput, 13  , sSnumber )
		.Parameters.Append .CreateParameter( "@Judge_Idx"  ,adInteger , adParamInput, 0   , Session("Judge_Idx") )
		.Parameters.Append .CreateParameter( "@sKind"      ,adVarChar , adParamInput, 3   , sKind )
		.Parameters.Append .CreateParameter( "@sClass"     ,adVarChar , adParamInput, 3   , sClass )
		.Parameters.Append .CreateParameter( "@sOnTime"  ,adVarChar , adParamInput, 2   , sOnTime )
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
ntpl.setTplDir( INTRA_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "application/leftMenu.html"
ntpl.setFile "MAIN", "application/applicationL.html"
ntpl.setFile "FOOTER", "_inc/footer.html"


call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(common_code_arrList(CCODE_Idx,iLoop))=sPcode,"admin_left_over","" ) ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST")
End If

call ntpl.setBlock("MAIN", array("APPLICATION_LOOP","LOOP_NODATA"))
'// BLOCK 부분 처리

If cntList > -1 Then 
	for iLoop = 0 to cntList

	StateTxt = ""

	If arrList(FI_State,iLoop) = "0" Then 
		StateTxt = "접수완료"
	ElseIf arrList(FI_State,iLoop) = "1" Then 
		StateTxt = "<font color=""#11179a"">입금대기</font>"
	ElseIf arrList(FI_State,iLoop) = "2" Then 
		StateTxt = "<font color=""#9a1134"">접수취소</font>"
	ElseIf arrList(FI_State,iLoop) = "3" Then 
		StateTxt = "<font color=""#9a1134"">불합격</font>"
	ElseIf arrList(FI_State,iLoop) = "4" Then 
		StateTxt = "<font color=""#9a1134"">미응시(불합격)</font>"
	ElseIf arrList(FI_State,iLoop) = "10" Then 
		StateTxt = "<font color=""#11179a"">합격</font>"
	End If

	PrograName = arrList(FI_ProgramNema,iLoop)

	If arrList(FI_Class,iLoop) = "1" Then
		PrograName = PrograName & " 1급"
	ElseIf arrList(FI_Class,iLoop) = "2" Then
		PrograName = PrograName & " 2급"
	ElseIf arrList(FI_Class,iLoop) = "3" Then
		PrograName = PrograName & " [SPECIAL]"
	End If

	If arrList(FI_Kind,iLoop) = "1" Then
		PrograName = PrograName & " [필기]"
	ElseIf arrList(FI_Kind,iLoop) = "2" Then
		PrograName = PrograName & " [실기]"	
	End If

		ntpl.setBlockReplace array( _
			 array("rownum" , arrList(FI_rownum,iLoop)  ) _
			,array("Idx" , arrList(FI_Idx,iLoop)  ) _
			,array("UserId", arrList(FI_UserId,iLoop) ) _
			,array("UserName", arrList(FI_UserName,iLoop) ) _
			,array("UserPhone", arrList(FI_UserHphone1,iLoop) &"-"& arrList(FI_UserHphone2,iLoop) &"-"& arrList(FI_UserHphone3,iLoop) ) _
			,array("ProgramNema" , PrograName  ) _
			,array("AreaName" , arrList(FI_AreaName,iLoop)  ) _
			,array("OnData" , arrList(FI_OnData,iLoop)  ) _
			,array("InData" , arrList(FI_InData,iLoop)  ) _
			,array("State" , StateTxt  ) _

			,array("Snumber" , IIF(arrList(FI_Snumber,iLoop)="","&nbsp;",arrList(FI_Snumber,iLoop))  ) _
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
	

	,array("pageNo"    , pageNo ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sOnDate" , sOnDate ) _
	,array("sPcode", sPcode ) _
	,array("sArea", sArea ) _
	,array("sId", sId ) _
	,array("sName", sName ) _
	,array("sPhone3", sPhone3 ) _
	,array("sState", sState ) _
	,array("sOnTime", sOnTime ) _

	,array("SearchOnDateOption", SearchOnDateOption ) _
	,array("sSnumber", sSnumber ) _
	,array("KindOption", KindOption ) _
	,array("ClassOption", ClassOption ) _
	,array("hoursOption", hoursOption ) _

	,array("s1Day"    , Date() ) _
	,array("s7Day"    , Date() -7 ) _
	,array("s30Day"   , Date() -30 ) _

	,array("leftMenuOverClass1"   , "admin_left_over" ) _
	,array("leftMenuOverClass2"   , "" ) _

	,array("Judge_Id", Session("Judge_Id") ) _
	,array("Judge_Name", Session("Judge_Name") ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>