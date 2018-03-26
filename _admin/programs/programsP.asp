<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim alertMsg      : alertMsg      = ""
Dim Idx           : Idx           = RequestSet("Idx"           , "POST" , "")
Dim actType       : actType       = RequestSet("actType"       , "POST" , "")
Dim CodeIdx       : CodeIdx       = RequestSet("CodeIdx"       , "POST" , 0 )
Dim OnData        : OnData        = RequestSet("OnData"        , "POST" , "")
Dim OnDataHours   : OnDataHours   = RequestSet("OnDataHours"   , "POST" , "")
Dim OnDataMinutes : OnDataMinutes = RequestSet("OnDataMinutes" , "POST" , "")
Dim areaIdx       : areaIdx       = RequestSet("areaIdx"       , "POST" , 0 )
Dim Pay           : Pay           = RequestSet("Pay"           , "POST" , 0 )
Dim CodeKind      : CodeKind      = RequestSet("CodeKind"      , "POST" , 0 )
Dim CodeClass     : CodeClass     = RequestSet("CodeClass"     , "POST" , 0 )
Dim StartDate     : StartDate     = RequestSet("StartDate"     , "POST" , "")
Dim EndDate       : EndDate       = RequestSet("EndDate"       , "POST" , "")
Dim MaxNumber     : MaxNumber     = RequestSet("MaxNumber"     , "POST" , 0 )
Dim pageNo        : pageNo        = RequestSet("pageNo"        , "POST" , 1 )
Dim sOnDate       : sOnDate       = RequestSet("sOnDate"       , "POST" , "")
Dim sPcode        : sPcode        = RequestSet("sPcode"        , "POST" , "")
Dim sName         : sName         = RequestSet("sName"         , "POST" , "")
Dim sKind         : sKind         = RequestSet("sKind"         , "POST" , "")
Dim sClass        : sClass        = RequestSet("sClass"        , "POST" , "")

OnData = OnData & " " & OnDataHours & ":" & OnDataMinutes


Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sName="      & sName &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass



Call Expires()
Call dbopen()
	If actType = "INSERT" Then 
		Call Insert()
		If FV_CNT > 0 Then 
			alertMsg = "중복된 프로그램명,시행일자 입니다. 정보를 확인해 주세요."
		else
			alertMsg = "입력되었습니다."
		End If		
	ElseIf actType = "UPDATE" Then 
		Call Update()
		If FV_CNT > 0 Then 
			alertMsg = "중복된 프로그램명,시행일자 입니다. 정보를 확인해 주세요."
		else
			alertMsg = "수정되었습니다."
		End If	
	ElseIf actType = "DELETE" Then 
		Call Delete()
		alertMsg = "삭제되었습니다."
	Else
		alertMsg = "[actType] 이 없습니다."
	End If
	
Call dbclose()


'입력
Sub Insert()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "set nocount on;" &_
	"DECLARE @areaIdx INT ,@CodeIdx INT,@OnData VARCHAR(20),@Pay INT,@CNT INT " &_
	"DECLARE @Kind INT,@Class INT,@StartDate VARCHAR(10),@EndDate VARCHAR(10),@MaxNumber INT " &_
	"SET @areaIdx   = ? " &_
	"SET @CodeIdx   = ? " &_
	"SET @OnData    = ? " &_
	"SET @Pay       = ? " &_
	"SET @Kind      = ? " &_
	"SET @Class     = ? " &_
	"SET @StartDate = ? " &_
	"SET @EndDate   = ? " &_
	"SET @MaxNumber = ? " &_

	"SET @CNT       = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM] A INNER JOIN [dbo].[SP_PROGRAM_ON_AREA] B ON(A.[Idx] = B.[ProgramIdx]) WHERE A.[CodeIdx] = @CodeIdx AND A.[OnData] = @OnData AND A.[Dellfg] = 0 AND B.[AreaIdx] =@areaIdx ) " &_
	
	"IF @CNT = 0 " &_
	"BEGIN " &_
	"	DECLARE @SCOPE_IDX INT " &_

	"	INSERT INTO [dbo].[SP_PROGRAM]( " &_
	"		 [CodeIdx] " &_
	"		,[OnData] " &_
	"		,[Dellfg] " &_
	"		,[InDate] " &_
	"		,[Pay] " &_
	"		,[StartDate] " &_
	"		,[EndDate] " &_
	"		,[MaxNumber] " &_
	"		,[Kind] " &_
	"		,[Class] " &_
	"	)VALUES( " &_
	"		 @CodeIdx " &_
	"		,@OnData " &_
	"		,0 " &_
	"		,getDate() " &_
	"		,@Pay " &_
	"		,@StartDate " &_
	"		,@EndDate " &_
	"		,@MaxNumber " &_
	"		,@Kind " &_
	"		,@Class " &_
	"	) " &_

	"	SET @SCOPE_IDX = SCOPE_IDENTITY() ;" &_

	"	INSERT INTO [dbo].[SP_PROGRAM_ON_AREA]( " &_
	"		 [ProgramIdx] " &_
	"		,[AreaIdx] " &_
	"	)VALUES( " &_
	"		 @SCOPE_IDX " &_
	"		,@areaIdx " &_
	"	) " &_
	"END " &_
	"SELECT @CNT AS [CNT] "
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@areaIdx"   ,adInteger , adParamInput, 0    , areaIdx )
		.Parameters.Append .CreateParameter( "@CodeIdx"   ,adInteger , adParamInput, 0    , CodeIdx )
		.Parameters.Append .CreateParameter( "@OnData"    ,adVarChar , adParamInput, 20   , OnData )
		.Parameters.Append .CreateParameter( "@Pay"       ,adInteger , adParamInput, 0    , Pay )

		.Parameters.Append .CreateParameter( "@Kind"      ,adInteger , adParamInput, 0    , CodeKind )
		.Parameters.Append .CreateParameter( "@Class"     ,adInteger , adParamInput, 0    , CodeClass )
		.Parameters.Append .CreateParameter( "@StartDate" ,adVarChar , adParamInput, 10   , StartDate )
		.Parameters.Append .CreateParameter( "@EndDate"   ,adVarChar , adParamInput, 10   , EndDate )
		.Parameters.Append .CreateParameter( "@MaxNumber" ,adInteger , adParamInput, 0    , MaxNumber )

		set objRs = .Execute
	End with
	call cmdclose()

	CALL setFieldValue(objRs, "FV")
	objRs.close	: Set objRs = Nothing
End Sub
'수정
Sub Update()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "set nocount on;" &_
	"DECLARE @IDX INT " &_
	"SET @IDX = ? " &_

	"DECLARE @areaIdx INT ,@CodeIdx INT,@OnData VARCHAR(20),@Pay INT,@CNT INT " &_
	"DECLARE @Kind INT,@Class INT,@StartDate VARCHAR(10),@EndDate VARCHAR(10),@MaxNumber INT " &_
	"SET @areaIdx   = ? " &_
	"SET @CodeIdx   = ? " &_
	"SET @OnData    = ? " &_
	"SET @Pay       = ? " &_
	"SET @Kind      = ? " &_
	"SET @Class     = ? " &_
	"SET @StartDate = ? " &_
	"SET @EndDate   = ? " &_
	"SET @MaxNumber = ? " &_

	"SET @CNT       = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM] A INNER JOIN [dbo].[SP_PROGRAM_ON_AREA] B ON(A.[Idx] = B.[ProgramIdx]) WHERE A.[CodeIdx] = @CodeIdx AND A.[OnData] = @OnData AND A.[Idx] != @IDX AND A.[Dellfg] = 0 AND B.[AreaIdx] =@areaIdx ) " &_	


	"IF @CNT = 0 " &_
	"BEGIN " &_
	
	"	UPDATE [dbo].[SP_PROGRAM] SET " &_
	"		 [CodeIdx]   = @CodeIdx " &_
	"		,[OnData]    = @OnData " &_
	"		,[Pay]       = @Pay " &_
	"		,[StartDate] = @StartDate " &_
	"		,[EndDate]   = @EndDate " &_
	"		,[MaxNumber] = @MaxNumber " &_
	"		,[Kind]      = @Kind " &_
	"		,[Class]     = @Class " &_
	"	WHERE [Idx] = @IDX " &_

	"	UPDATE [dbo].[SP_PROGRAM_ON_AREA] SET " &_
	"		 [AreaIdx]   = @areaIdx " &_
	"	WHERE [ProgramIdx] = @IDX " &_

	"END " &_
	"SELECT @CNT AS [CNT] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"       ,adInteger , adParamInput, 0    , Idx )
		.Parameters.Append .CreateParameter( "@areaIdx"   ,adInteger , adParamInput, 0    , areaIdx )
		.Parameters.Append .CreateParameter( "@CodeIdx"   ,adInteger , adParamInput, 0    , CodeIdx )
		.Parameters.Append .CreateParameter( "@OnData"    ,adVarChar , adParamInput, 20   , OnData )
		.Parameters.Append .CreateParameter( "@Pay"       ,adInteger , adParamInput, 0    , Pay )

		.Parameters.Append .CreateParameter( "@Kind"      ,adInteger , adParamInput, 0    , CodeKind )
		.Parameters.Append .CreateParameter( "@Class"     ,adInteger , adParamInput, 0    , CodeClass )
		.Parameters.Append .CreateParameter( "@StartDate" ,adVarChar , adParamInput, 10   , StartDate )
		.Parameters.Append .CreateParameter( "@EndDate"   ,adVarChar , adParamInput, 10   , EndDate )
		.Parameters.Append .CreateParameter( "@MaxNumber" ,adInteger , adParamInput, 0    , MaxNumber )
		set objRs = .Execute
	End with
	call cmdclose()

	CALL setFieldValue(objRs, "FV")
	objRs.close	: Set objRs = Nothing
End Sub
'삭제
Sub Delete()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "DECLARE @S VARCHAR (max) " &_
	"DECLARE @T TABLE(T_INT INT) " &_
	"SET @S = ? " &_
	"WHILE CHARINDEX(',',@S)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(T_INT) VALUES( SUBSTRING(@S,1,CHARINDEX(',',@S)-1) ) " &_
	"	SET @S=SUBSTRING(@S,CHARINDEX(',',@S)+1,LEN(@S))  " &_
	"END " &_
	"IF CHARINDEX(',',@S)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(T_INT) VALUES( SUBSTRING(@S,1,LEN(@S)) ) " &_
	"END " &_
	
	
	"UPDATE [dbo].[SP_PROGRAM] SET " &_
	"	[Dellfg] = 1 " &_
	"WHERE [Idx] in( SELECT T_INT FROM @T ) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adVarChar , adParamInput, 8000 , Idx )
		.Execute
	End with
	call cmdclose()
End Sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN">
<html>
<head>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
</head>
<script language=javascript>
	if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
	top.location.href = "programsL.asp?<%=PageParams%>";
</script>
</html>