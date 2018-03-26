<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim Msg    : Msg    = ""
Dim idx    : idx    = RequestSet("idx"    , "POST" , "" )
Dim values : values = RequestSet("values" , "POST" , "" )

Call Expires()

If Session("Admin_Idx") = "" Then 
	Msg = "login,"
	Response.write Msg
	Response.End
End If

If idx = "" And values = "" Then 
	Msg = "fail,"
	Response.write Msg
	Response.End
End If

Call dbopen()

	Call getView()

	'수검번호 생성
	'응시년도2자리 + 응시월2자리 + 검정장3자리 + 검정과목1자리 + 필기/실기1자리 + 급수 1자리 + 등록번호3자리
	Dim sNumber1 : sNumber1 = Mid(FI_OnData,3,2)
	Dim sNumber2 : sNumber2 = Mid(FI_OnData,6,2)
	Dim sNumber3 : sNumber3 = FI_AreaCode
	Dim sNumber4 : sNumber4 = FI_ProgramCode
	Dim sNumber5 : sNumber5 = FI_Kind
	Dim sNumber6 : sNumber6 = FI_Class
	'Dim sNumber7 : sNumber7 = lpad( FI_AppCode , "0" , 3 )

	Dim sNumber : sNumber = sNumber1 & sNumber2 & sNumber3 & sNumber4 & sNumber5 & sNumber6

	Call Update()
	Msg = "ok," & IIF(RESULT_SnumberCheck = 1 , RESULT_Snumber , "")

Call dbclose()

Response.write Msg
Response.End



'수검번호 정보 검색
Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @AppIdx INT , @ProgramIdx INT , @AreaIdx INT ;" &_
	"SET @AppIdx = ? " &_
	
	"SELECT " &_
	"	 @ProgramIdx = [ProgramIdx] " &_
	"	,@AreaIdx    = [AreaIdx] " &_
	"FROM [dbo].[SP_PROGRAM_APP] " &_
	"WHERE [Idx] = @AppIdx " &_

	"SELECT " &_
	"	 convert(varchar(10),[OnData],23) AS [OnData]" &_
	"	,@ProgramIdx AS [ProgramIdx] " &_
	"	,@AreaIdx AS [AreaIdx] " &_
	"	,( SELECT COUNT(*) + 1 FROM [dbo].[SP_COMM_CODE2] where [PIdx] = 17 and [idx] < [CodeIdx] ) AS [ProgramCode] " &_
	"	,( SELECT ISNULL([Code],'000') FROM [dbo].[SP_PROGRAM_AREA] where [Idx] = @AreaIdx ) AS [AreaCode] " &_
	"	,( SELECT COUNT(*) + 1 FROM [dbo].[SP_PROGRAM_APP] where [ProgramIdx] = @ProgramIdx AND [AreaIdx] = @AreaIdx AND [Idx] < @AppIdx ) AS [AppCode] " &_
	"	,ISNULL([Kind],0) AS [Kind] " &_
	"	,ISNULL([Class],0) AS [Class] " &_
	"FROM [dbo].[SP_PROGRAM]" &_
	"WHERE [Idx] =  @ProgramIdx "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@AppIdx" ,adInteger , adParamInput , 0 , idx )
		set objRs = .Execute
	End with
	call cmdclose()
	'프로그램정보
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub


'수정
Sub Update()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_
	"declare @State int,@Idx int,@Snumber VARCHAR(50),@appCount varchar(3),@SnumberCheck INT,@ProgramIdx int , @AreaIdx int,@ip VARCHAR(50) ;" &_
	"set @State   = ? " &_
	"set @Idx     = ? " &_
	"set @Snumber = ? " &_
	"set @ProgramIdx = ? " &_
	"set @AreaIdx = ? " &_
	"set @ip = ? " &_
	"set @appCount   = ( SELECT COUNT(*) + 1 FROM [dbo].[SP_PROGRAM_APP] where [ProgramIdx] = @ProgramIdx AND [AreaIdx] = @AreaIdx AND [Idx] < @Idx ) " &_

	"set @SnumberCheck = 0 " &_

	"IF @State = 0 " &_
	"BEGIN " &_
	"	UPDATE [dbo].[SP_PROGRAM_APP] SET " &_
	"		 [Snumber] = @Snumber + (REPLICATE('0', 3-LEN(@appCount)) + @appCount) " &_
	"		,[NocachDate] = GETDATE() " &_
	"	WHERE [Idx] = @Idx AND [PayMode]='SC0040' AND ( [Snumber] IS NULL OR [Snumber] = '' ) " &_
	"	SET @SnumberCheck = 1 " &_
	"END " &_

	"UPDATE [dbo].[SP_PROGRAM_APP] SET " &_
	"	 [State] = @State " &_
	"WHERE [Idx] = @Idx " &_

	"INSERT INTO [dbo].[SP_PROGRAM_APP_log]( " &_
	"	 [Idx]" &_
	"	,[State]" &_
	"	,[InData]" &_
	"	,[ip]" &_
	")VALUES( "  &_
	"	 @Idx" &_
	"	,@State" &_
	"	,GETDATE()" &_
	"	,@ip" &_
	")" &_

	"SELECT @SnumberCheck AS [SnumberCheck] , (@Snumber + (REPLICATE('0', 3-LEN(@appCount)) + @appCount) )  AS [Snumber] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@State"      ,adInteger , adParamInput , 0  , values )
		.Parameters.Append .CreateParameter( "@Idx"        ,adInteger , adParamInput , 0  , idx )
		.Parameters.Append .CreateParameter( "@Snumber"    ,adVarChar , adParamInput , 50 , sNumber )
		.Parameters.Append .CreateParameter( "@ProgramIdx" ,adInteger , adParamInput , 0  , FI_ProgramIdx )
		.Parameters.Append .CreateParameter( "@AreaIdx"    ,adInteger , adParamInput , 0  , FI_AreaIdx )
		.Parameters.Append .CreateParameter( "@ip"         ,adVarChar , adParamInput , 50 , g_uip )
		
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "RESULT")
	Set objRs = Nothing
End Sub
%>