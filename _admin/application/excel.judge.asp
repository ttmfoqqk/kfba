<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1

Dim pageNo     : pageNo    = RequestSet("pageNo" , "GET" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate" , "GET" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate", "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "")
Dim sState     : sState     = RequestSet("sState"    , "GET" , "")
Dim sId        : sId        = RequestSet("sId"    , "GET" , "")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"    , "GET" , "")
Dim sBirth     : sBirth     = RequestSet("sBirth"    , "GET" , "")

Call Expires()
Call dbopen()
	Call GetList()
Call dbclose()


Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
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
	"		,A.[InData] " &_
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
	"ORDER BY rownum DESC "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
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



"<Worksheet ss:Name=""심사위원 응시 리스트""> " &_
"<Table> " &_
"	<Column ss:Width='30'/> " &_
"	<Column ss:Width='80'/> " &_
"	<Column ss:Width='80'/> " &_
"	<Column ss:Width='100'/> " &_
"	<Column ss:Width='200'/> " &_
"	<Column ss:Width='80'/> " &_
"	<Column ss:Width='150'/> " &_
"	<Row> "&_
"		<Cell><Data ss:Type=""String"">NO</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">ID</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">이름</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">연락처</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">자격종목</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">접수현황</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">접수일자</Data></Cell> "&_
"	</Row> "

If cntList > -1 Then 
	for iLoop = 0 to cntList

		StateTxt = ""
		ss_Color = ""

		If arrList(FI_State,iLoop) = "1" Then 
			StateTxt = "접수"
		ElseIf arrList(FI_State,iLoop) = "0" Then 
			StateTxt = "승인"
			ss_Color = " ss:StyleID='sBlue'"
		ElseIf arrList(FI_State,iLoop) = "2" Then 
			StateTxt = "불합격"
			ss_Color = " ss:StyleID='sRed'"
		End If

		PrograName = arrList(FI_ProgramNema,iLoop)

		If arrList(FI_ProgramKind,iLoop) = "1" Then
			PrograName = PrograName & " [필기]"
		ElseIf arrList(FI_ProgramKind,iLoop) = "2" Then
			PrograName = PrograName & " [실기]"
		ElseIf arrList(FI_ProgramKind,iLoop) = "3" Then
			PrograName = PrograName & " [SPECIAL]"
		End If
		
		tmp_html = tmp_html & "" &_
		"	<Row> "&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_rownum,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserId,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserName,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserHphone1,iLoop) &"-"& arrList(FI_UserHphone2,iLoop) &"-"& arrList(FI_UserHphone3,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & PrograName & "</Data></Cell>"&_
		"		<Cell"&ss_Color&"><Data ss:Type=""String"">" & StateTxt & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_InData,iLoop) & "</Data></Cell>"&_
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
Response.AddHeader "Content-Disposition","attachment; filename=심사위원 응시 리스트 " & Now() & ".xls"
%>