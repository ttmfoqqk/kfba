<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1

Dim pageNo   : pageNo    = RequestSet("pageNo" , "GET" , 1)

Dim sIndate  : sIndate   = RequestSet("sIndate" , "GET" , "")
Dim sOutdate : sOutdate  = RequestSet("sOutdate", "GET" , "")
Dim sName    : sName     = RequestSet("sName"   , "GET" , "")
Dim sAddr    : sAddr     = RequestSet("sAddr"   , "GET" , "")
Dim sCode    : sCode     = RequestSet("sCode"   , "GET" , "")
Dim sTel     : sTel      = RequestSet("sTel"    , "GET" , "")
Dim sPcode   : sPcode    = RequestSet("sPcode"  , "GET" , "")

Call Expires()
Call dbopen()
	Call GetList()
Call dbclose()

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_

	"DECLARE @sIndate VARCHAR(10) , @sOutdate VARCHAR(10) , @sName VARCHAR(200) , @sAddr VARCHAR(200) , @sCode VARCHAR(3) , @sTel VARCHAR(50) , @sPcode VARCHAR(10) ;" &_
	"SET @sIndate  = ?; " &_
	"SET @sOutdate = ?; " &_
	"SET @sName    = ?; " &_
	"SET @sAddr    = ?; " &_
	"SET @sCode    = ?; " &_
	"SET @sTel     = ?; " &_
	"SET @sPcode   = ?; " &_

	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over (order by [Idx] ) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,[Idx] " &_
	"		,[Name] " &_
	"		,[Addr] " &_
	"		,[Tel] " &_
	"		,[Code] " &_
	"		,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [AddrIdx] ) AS [AddrCode]" &_
	"	FROM [dbo].[SP_PROGRAM_AREA] " &_

	"	WHERE [Dellfg] = 0 " &_
	"	AND CASE @sName WHEN '' THEN '' ELSE [Name] END LIKE '%'+@sName+'%' "&_
	"	AND CASE @sAddr WHEN '' THEN '' ELSE [Addr] END LIKE '%'+@sAddr+'%' " &_
	"	AND CASE @sCode WHEN '' THEN '' ELSE [Code] END = @sCode " &_
	"	AND CASE @sTel WHEN '' THEN '' ELSE [Tel] END LIKE '%'+@sTel+'%' " &_
	"	AND CASE @sIndate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[Indate],23) END >= @sIndate " &_
	"	AND CASE @sOutdate WHEN '' THEN '' ELSE CONVERT(VARCHAR,[Indate],23) END <= @sOutdate " &_
	"	AND CASE @sPcode WHEN '' THEN '' ELSE [CodeIdx] END = @sPcode " &_

	") " &_
	"SELECT L.* " &_
	"FROM LIST L " &_
	"ORDER BY rownum DESC "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@sIndate"  ,adVarChar , adParamInput, 10  , sIndate )
		.Parameters.Append .CreateParameter( "@sOutdate" ,adVarChar , adParamInput, 10  , sOutdate )
		.Parameters.Append .CreateParameter( "@sName"    ,adVarChar , adParamInput, 200 , sName )
		.Parameters.Append .CreateParameter( "@sAddr"    ,adVarChar , adParamInput, 200 , sAddr )
		.Parameters.Append .CreateParameter( "@sCode"    ,adVarChar , adParamInput, 3   , sCode )
		.Parameters.Append .CreateParameter( "@sTel"     ,adVarChar , adParamInput, 50  , sTel )
		.Parameters.Append .CreateParameter( "@sPcode"   ,adVarChar , adParamInput, 10  , sPcode )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
	End If
	objRs.close	: Set objRs = Nothing
End Sub

Dim tmp_html : tmp_html = "" &_
"<?xml version=""1.0"" encoding=""EUC-KR""?>" &_
"<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">" &_

"<Styles> " &_
"  <Style ss:ID='Default' ss:Name='Normal'> " &_
"  <Alignment ss:Vertical='Center'/> " &_
"  <Borders/> " &_
"  <Font ss:FontName='굴림' x:CharSet='129' x:Family='Modern' ss:Size='10'/> " &_
"  <Interior/> " &_
"  <NumberFormat/> " &_
"  <Protection/> " &_
"  </Style> " &_
"  <Style ss:ID='s30'> " &_
"  <Alignment ss:Horizontal='Center' ss:Vertical='Center' ss:WrapText='1'/> " &_
"  <Borders> " &_
"    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"  </Borders> " &_
"  <Interior ss:Color='#FFFF99' ss:Pattern='Solid'/> " &_
"  <NumberFormat ss:Format='@'/> " &_
"  </Style> " &_
"  <Style ss:ID='s31'> " &_
"  <Borders> " &_
"    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/> " &_
"  </Borders> " &_
"  <Interior ss:Color='#FFFF99' ss:Pattern='Solid'/> " &_
"  </Style> " &_
" </Styles> " &_

"<Worksheet ss:Name=""검정장리스트""> " &_
"<Table> " &_
"	<Column ss:Width='30'/> " &_
"	<Column ss:Width='30'/> " &_
"	<Column ss:Width='60'/> " &_
"	<Column ss:Width='300'/> " &_
"	<Column ss:Width='300'/> " &_
"	<Column ss:Width='100'/> " &_
"	<Row> "&_
"		<Cell><Data ss:Type=""String"">NO</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">코드</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">지역</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">검정장 이름</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">주소</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">연락처</Data></Cell> "&_
"	</Row> "

If cntList > -1 Then 
	for iLoop = 0 to cntList
		
		tmp_html = tmp_html & "" &_
		"	<Row> "&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_rownum,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_Code,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_AddrCode,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_Name,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_Addr,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_Tel,iLoop) & "</Data></Cell>"&_
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
Response.AddHeader "Content-Disposition","attachment; filename=검정장리스트 " & Now() & ".xls"

%>