<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim alertMsg : alertMsg   = ""
Dim actType  : actType    = RequestSet("actType"  , "POST" , "" )
Dim Idx      : Idx        = RequestSet("Idx"      , "POST" , "" )
Dim State    : State      = RequestSet("State"    , "POST" , 0 )
Dim Bigo     : Bigo       = TagEncode( RequestSet("Bigo"     , "POST" , "" ) )

Dim pageNo   : pageNo     = RequestSet("pageNo"   , "POST" , 1)
Dim sIndate  : sIndate    = RequestSet("sIndate"  , "POST" , "")
Dim sOutdate : sOutdate   = RequestSet("sOutdate" , "POST" , "")
Dim sOnDate  : sOnDate    = RequestSet("sOnDate"  , "POST" , "")
Dim sPcode   : sPcode     = RequestSet("sPcode"   , "POST" , "")
Dim sArea    : sArea      = RequestSet("sArea"    , "POST" , "")
Dim sId      : sId        = RequestSet("sId"      , "POST" , "")
Dim sName    : sName      = RequestSet("sName"    , "POST" , "")
Dim sPhone3  : sPhone3    = RequestSet("sPhone3"  , "POST" , "")
Dim sState   : sState     = RequestSet("sState"   , "POST" , "")
Dim sSnumber : sSnumber   = RequestSet("sSnumber" , "POST" , "")
Dim sKind    : sKind      = RequestSet("sKind"    , "POST" , "")
Dim sClass   : sClass     = RequestSet("sClass"   , "POST" , "")

Dim sOnTime  : sOnTime    = RequestSet("sOnTime"  , "POST" , "")


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


Call Expires()
Call dbopen()
	
	If actType = "UPDATE" Then 

		Call getView()

		'���˹�ȣ ����
		'���ó⵵2�ڸ� + ���ÿ�2�ڸ� + ������3�ڸ� + ��������1�ڸ� + �ʱ�/�Ǳ�1�ڸ� + �޼� 1�ڸ� + ��Ϲ�ȣ3�ڸ�
		Dim sNumber1 : sNumber1 = Mid(FI_OnData,3,2)
		Dim sNumber2 : sNumber2 = Mid(FI_OnData,6,2)
		Dim sNumber3 : sNumber3 = FI_AreaCode
		Dim sNumber4 : sNumber4 = FI_ProgramCode
		Dim sNumber5 : sNumber5 = FI_Kind
		Dim sNumber6 : sNumber6 = FI_Class
		'Dim sNumber7 : sNumber7 = lpad( FI_AppCode , "0" , 3 )

		Dim sNumber : sNumber = sNumber1 & sNumber2 & sNumber3 & sNumber4 & sNumber5 & sNumber6

		Call Update()
		alertMsg = "�����Ǿ����ϴ�."
	ElseIf actType = "DELETE" Then 
		Call Delete()
		alertMsg = "�����Ǿ����ϴ�."
	Else
		alertMsg = "[actType] �� �����ϴ�."
	End If
	
Call dbclose()

'���˹�ȣ ���� �˻�
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
		.Parameters.Append .CreateParameter( "@AppIdx" ,adInteger , adParamInput , 0 , Idx )
		set objRs = .Execute
	End with
	call cmdclose()
	'���α׷�����
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub


'����
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_
	"declare @State int,@Idx int,@Snumber VARCHAR(50),@appCount varchar(3),@ProgramIdx int , @AreaIdx int, @Bigo VARCHAR(MAX) , @ip VARCHAR(50) ;" &_
	"set @State   = ? " &_
	"set @Idx     = ? " &_
	"set @Snumber = ? " &_
	"set @Bigo    = ? " &_
	"set @ProgramIdx = ? " &_
	"set @AreaIdx = ? " &_
	"set @ip = ? " &_
	"set @appCount   = ( SELECT COUNT(*) + 1 FROM [dbo].[SP_PROGRAM_APP] where [ProgramIdx] = @ProgramIdx AND [AreaIdx] = @AreaIdx AND [Idx] < @Idx ) " &_

	"IF @State = 0 " &_
	"BEGIN " &_
	"	UPDATE [dbo].[SP_PROGRAM_APP] SET " &_
	"		 [Snumber] = @Snumber + (REPLICATE('0', 3-LEN(@appCount)) + @appCount) " &_
	"		,[NocachDate] = GETDATE() " &_
	"	WHERE [Idx] = @Idx AND [PayMode]='SC0040' AND ( [Snumber] IS NULL OR [Snumber] = '' ) " &_
	"END " &_

	"UPDATE [dbo].[SP_PROGRAM_APP] SET " &_
	"	 [State] = @State " &_
	"	,[Bigo]  = @Bigo " &_
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
	")"

	

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@State"      ,adInteger     , adParamInput , 0         , State )
		.Parameters.Append .CreateParameter( "@Idx"        ,adInteger     , adParamInput , 0         , Idx )
		.Parameters.Append .CreateParameter( "@Snumber"    ,adVarChar     , adParamInput , 50        , sNumber )
		.Parameters.Append .CreateParameter( "@Bigo"       ,adLongVarChar , adParamInput , 2147483647, Bigo )
		.Parameters.Append .CreateParameter( "@ProgramIdx" ,adInteger     , adParamInput , 0         , FI_ProgramIdx )
		.Parameters.Append .CreateParameter( "@AreaIdx"    ,adInteger     , adParamInput , 0         , FI_AreaIdx )
		.Parameters.Append .CreateParameter( "@ip"         ,adVarChar     , adParamInput , 50        , g_uip )
		.Execute
	End with
	call cmdclose()
End Sub
'����
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
	
	
	"UPDATE [dbo].[SP_PROGRAM_AREA] SET " &_
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
	top.location.href = "applicationL.asp?<%=PageParams%>";
</script>
</html>