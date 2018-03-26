<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1

Dim pageNo     : pageNo     = RequestSet("pageNo"   , "GET" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate"  , "GET" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate" , "GET" , "")
Dim sOnDate    : sOnDate    = RequestSet("sOnDate"  , "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "")
Dim sArea      : sArea      = RequestSet("sArea"    , "GET" , "")

Dim sId        : sId        = RequestSet("sId"      , "GET" , "")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"  , "GET" , "")
Dim sState     : sState     = RequestSet("sState"   , "GET" , "")
Dim sSnumber   : sSnumber   = RequestSet("sSnumber" , "GET" , "")
Dim sKind      : sKind      = RequestSet("sKind"    , "GET" , "")
Dim sClass     : sClass     = RequestSet("sClass"   , "GET" , "")

Dim sOnTime    : sOnTime    = RequestSet("sOnTime"  , "GET" , "")


Call Expires()
Call dbopen()
	Call GetList()
Call dbclose()


Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @sIndate VARCHAR(10) , @sOutdate VARCHAR(10) , @sOnDate VARCHAR(10);" &_
	"DECLARE @sPcode VARCHAR(10) , @sArea VARCHAR(200) , @sId VARCHAR(50) , @sName VARCHAR(50) , @sPhone3 VARCHAR(4) , @sState VARCHAR(3) , @sSnumber VARCHAR(13) , @sKind VARCHAR(5) , @sClass VARCHAR(5),@sOnTime VARCHAR(2)  ;" &_
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
	"		,A.[InData] " &_
	"		,B.[UserId]" &_
	"		,B.[UserName]" &_
	"		,B.[UserBirth] " &_
	"		,B.[UserHphone1] " &_
	"		,B.[UserHphone2] " &_
	"		,B.[UserHphone3] " &_
	"		,B.[FirstName] " &_
	"		,B.[LastName] " &_
	"		,C.[Name] AS [ProgramNema]" &_
	"		,C.[Kind]" &_
	"		,C.[Class]" &_
	"		,C.[OnData] " &_
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
	"	AND ( CASE @sName WHEN '' THEN '' ELSE  B.[UserName] END LIKE '%'+@sName+'%' OR CASE @sName WHEN '' THEN '' ELSE B.[FirstName] END LIKE '%'+@sName+'%' )  " &_
	"	AND CASE @sPhone3 WHEN '' THEN '' ELSE B.[UserHphone3] END LIKE '%'+@sPhone3+'%' "&_
	"	AND CASE @sState WHEN '' THEN '' ELSE A.[State] END = @sState " &_
	"	AND CASE @sSnumber WHEN '' THEN '' ELSE A.[Snumber] END = @sSnumber " &_
	"	AND CASE @sKind WHEN '' THEN '' ELSE C.[Kind] END = @sKind " &_
	"	AND CASE @sClass WHEN '' THEN '' ELSE C.[Class] END = @sClass " &_

	"	AND CASE @sIndate WHEN '' THEN '' ELSE CONVERT(VARCHAR,A.[InData],23) END >= @sIndate " &_
	"	AND CASE @sOutdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,A.[InData],23) END <= @sOutdate " &_
	"	AND CASE @sOnDate WHEN '' THEN '' ELSE CONVERT(VARCHAR,C.[OnData],23) END = @sOnDate " &_
	"	AND CASE @sOnTime WHEN '' THEN '' ELSE CONVERT(VARCHAR(2),C.[OnData],108) END = @sOnTime " &_
	") " &_
	"SELECT L.* " &_
	"FROM LIST L " &_
	"ORDER BY rownum DESC; "


	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@sIndate"  ,adVarChar , adParamInput, 10  , sIndate )
		.Parameters.Append .CreateParameter( "@sOutdate" ,adVarChar , adParamInput, 10  , sOutdate )
		.Parameters.Append .CreateParameter( "@sOnDate"  ,adVarChar , adParamInput, 10  , sOnDate )
		.Parameters.Append .CreateParameter( "@sPcode"   ,adVarChar , adParamInput, 10  , sPcode )
		.Parameters.Append .CreateParameter( "@sArea"    ,adVarChar , adParamInput, 200 , sArea )
		.Parameters.Append .CreateParameter( "@sId"      ,adVarChar , adParamInput, 50  , sId )
		.Parameters.Append .CreateParameter( "@sName"    ,adVarChar , adParamInput, 50  , sName )
		.Parameters.Append .CreateParameter( "@sPhone3"  ,adVarChar , adParamInput, 4   , sPhone3 )
		.Parameters.Append .CreateParameter( "@sState"   ,adVarChar , adParamInput, 3   , sState )
		.Parameters.Append .CreateParameter( "@sSnumber" ,adVarChar , adParamInput, 13  , sSnumber )
		.Parameters.Append .CreateParameter( "@sKind"    ,adVarChar , adParamInput, 3   , sKind )
		.Parameters.Append .CreateParameter( "@sClass"   ,adVarChar , adParamInput, 3   , sClass )
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
	objRs.close	: Set objRs = Nothing
End Sub

Dim tmp_html : tmp_html = "" &_
"<?xml version=""1.0"" encoding=""EUC-KR""?>" &_
"<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">" &_

"<Styles> " &_
"  <Style ss:ID='sRed'> " &_
"  <Font ss:FontName=""굴림"" x:CharSet=""129"" x:Family=""Modern"" ss:Size=""10"" ss:Color=""#FF0000""/> " &_
"  </Style> " &_

"  <Style ss:ID='sBlue'> " &_
"  <Font ss:FontName=""굴림"" x:CharSet=""129"" x:Family=""Modern"" ss:Size=""10"" ss:Color=""#0000FF""/> " &_
"  </Style> " &_
" </Styles> " &_



"<Worksheet ss:Name=""검정 응시 리스트""> " &_
"<Table> " &_
"	<Column ss:Width='30'/> " &_
"	<Column ss:Width='80'/> " &_
"	<Column ss:Width='80'/> " &_
"	<Column ss:Width='150'/> " &_
"	<Column ss:Width='100'/> " &_
"	<Column ss:Width='200'/> " &_
"	<Column ss:Width='250'/> " &_
"	<Column ss:Width='150'/> " &_
"	<Column ss:Width='150'/> " &_
"	<Column ss:Width='100'/> " &_
"	<Column ss:Width='100'/> " &_
"	<Row> "&_
"		<Cell><Data ss:Type=""String"">NO</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">이름</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">영문이름</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">생년월일</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">연락처</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">자격종목</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">지정검정장</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">검정시행일</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">접수일자</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">수검번호</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">접수현황</Data></Cell> "&_
"	</Row> "

If cntList > -1 Then 
	for iLoop = 0 to cntList

		StateTxt = ""
		ss_Color = ""

		If arrList(FI_State,iLoop) = "0" Then 
			StateTxt = "접수완료"
		ElseIf arrList(FI_State,iLoop) = "1" Then 
			StateTxt = "입금대기"
			ss_Color = " ss:StyleID='sBlue'"
		ElseIf arrList(FI_State,iLoop) = "2" Then 
			StateTxt = "접수취소"
			ss_Color = " ss:StyleID='sRed'"
		ElseIf arrList(FI_State,iLoop) = "3" Then 
			StateTxt = "불합격"
			ss_Color = " ss:StyleID='sRed'"
		ElseIf arrList(FI_State,iLoop) = "4" Then 
			StateTxt = "미응시(불합격)"
			ss_Color = " ss:StyleID='sRed'"
		ElseIf arrList(FI_State,iLoop) = "10" Then 
			StateTxt = "합격"
			ss_Color = " ss:StyleID='sBlue'"
		End If

		PrograName = arrList(FI_ProgramNema,iLoop)

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
		
		tmp_html = tmp_html & "" &_
		"	<Row> "&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_rownum,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserName,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_LastName,iLoop) & " " & arrList(FI_FirstName,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserBirth,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserHphone1,iLoop) &"-"& arrList(FI_UserHphone2,iLoop) &"-"& arrList(FI_UserHphone3,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & PrograName & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_AreaName,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_OnData,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_InData,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_Snumber,iLoop) & "</Data></Cell>"&_
		"		<Cell"&ss_Color&"><Data ss:Type=""String"">" & StateTxt & "</Data></Cell>"&_
		"	</Row> "
	Next
Else
	tmp_html = tmp_html & "<Row><Cell><Data ss:Type=""String"">등록된 내용이 없습니다.</Data></Cell></Row>"
End If

tmp_html = tmp_html & "</Table></Worksheet></Workbook>"


Response.write tmp_html


Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=검정 응시 리스트 " & Now() & ".xls"
%>