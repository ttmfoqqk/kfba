<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList   : cntList   = -1


Dim pageNo     : pageNo     = RequestSet("pageNo","GET",1)
Dim sIndate    : sIndate    = RequestSet("sIndate","GET","")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate","GET","")
Dim sUserId    : sUserId    = RequestSet("sUserId","GET","")
Dim sUserName  : sUserName  = RequestSet("sUserName","GET","")
Dim sHphone3   : sHphone3   = RequestSet("sHphone3","GET","")
Dim sUserBirth : sUserBirth = RequestSet("sUserBirth","GET","")
Dim sState     : sState     = RequestSet("sState","GET","")

Call Expires()
Call dbopen()
	Call GetList()
Call dbclose()

Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
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
	"ORDER BY rownum desc "

	call cmdopen()
	with objCmd
		.CommandText = SQL
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

"<Worksheet ss:Name=""회원리스트""> " &_
"<Table> " &_
"	<Column ss:Width='30'/> " &_
"	<Column ss:Width='80'/> " &_
"	<Column ss:Width='80'/> " &_
"	<Column ss:Width='100'/> " &_
"	<Column ss:Width='100'/> " &_
"	<Column ss:Width='200'/> " &_
"	<Column ss:Width='150'/> " &_
"	<Column ss:Width='50'/> " &_
"	<Row> "&_
"		<Cell><Data ss:Type=""String"">NO</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">ID</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">이름</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">연락처</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">생년월일</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">이메일</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">가입일자</Data></Cell> "&_
"		<Cell><Data ss:Type=""String"">탈퇴여부</Data></Cell> "&_
"	</Row> "

If cntList > -1 Then 
	for iLoop = 0 to cntList
		
		tmp_html = tmp_html & "" &_
		"	<Row> "&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_rownum,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserId,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserName,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserHphone1,iLoop) &"-"& arrList(FI_UserHphone2,iLoop) &"-"& arrList(FI_UserHphone3,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & Mid(arrList(FI_UserBirth,iLoop),1,4) &"-"&Mid(arrList(FI_UserBirth,iLoop),5,2) &"-"&Mid(arrList(FI_UserBirth,iLoop),7,2) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserEmail,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & arrList(FI_UserIndate,iLoop) & "</Data></Cell>"&_
		"		<Cell><Data ss:Type=""String"">" & IIF(arrList(FI_UserDelFg,iLoop) = 0 , "사용" , "탈퇴" ) & "</Data></Cell>"&_
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
Response.AddHeader "Content-Disposition","attachment; filename=회원리스트 " & Now() & ".xls"

%>