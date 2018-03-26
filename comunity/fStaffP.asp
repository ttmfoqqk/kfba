<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%

Dim alertMsg     : alertMsg = ""
Dim actType      : actType     = RequestSet("actType" ,"POST","")

Dim Idx          : Idx         = RequestSet("Idx" ,"POST",0)
Dim OwnerName    : OwnerName   = TagEncode( RequestSet("OwnerName" ,"POST","") )
Dim ManagerName  : ManagerName = TagEncode( RequestSet("ManagerName" ,"POST","") )
Dim HomePage     : HomePage    = RequestSet("HomePage" ,"POST","")
Dim CompanyName  : CompanyName = TagEncode( RequestSet("CompanyName" ,"POST","") )
Dim Addr         : Addr        = TagEncode( RequestSet("Addr" ,"POST","") )
Dim Tel          : Tel         = TagEncode( RequestSet("Tel" ,"POST","") )
Dim Fax          : Fax         = TagEncode( RequestSet("Fax" ,"POST","") )
Dim Email        : Email       = TagEncode( RequestSet("Email" ,"POST","") )
Dim Title        : Title       = TagEncode( RequestSet("Title" ,"POST","") )
Dim Form         : Form        = RequestSet("Form" ,"POST",0)
Dim Kind         : Kind        = RequestSet("Kind" ,"POST",0)
Dim WorkArea     : WorkArea    = TagEncode( RequestSet("WorkArea" ,"POST","") )
Dim WorkTime     : WorkTime    = TagEncode( RequestSet("WorkTime" ,"POST","") )
Dim StaffCnt     : StaffCnt    = RequestSet("StaffCnt" ,"POST",0)
Dim Qualify      : Qualify     = TagEncode( RequestSet("Qualify" ,"POST","") )
Dim Files        : Files       = TagEncode( RequestSet("Files" ,"POST","") )
Dim Dates        : Dates       = TagEncode( RequestSet("Dates" ,"POST","") )
Dim Method       : Method      = TagEncode( RequestSet("Method" ,"POST","") )
Dim Pay          : Pay         = TagEncode( RequestSet("Pay" ,"POST","") )
Dim insure       : insure      = RequestSet("insure" ,"POST",0)
Dim Bigo         : Bigo        = TagEncode( RequestSet("Bigo" ,"POST","") )
Dim UserIdx      : UserIdx     = IIF( session("UserIdx")="",0,session("UserIdx") )
Dim Pwd          : Pwd         = RequestSet("Pwd" ,"POST","")

Dim pageNo       : pageNo      = RequestSet("pageNo" ,"POST","")
Dim sName        : sName       = RequestSet("sName" ,"GET",0)
Dim sId          : sId         = RequestSet("sId" ,"GET",0)
Dim sTitle       : sTitle      = RequestSet("sTitle" ,"GET",0)
Dim sContant     : sContant    = RequestSet("sContant" ,"GET",0)
Dim sWord        : sWord       = RequestSet("sWord" ,"GET","")

HomePage = TagEncode( Replace( Replace( LCase(HomePage) ,"https://","")  ,"http://","") )

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&BoardKey=" & BoardKey &_
		"&sName="    & sName &_
		"&sId="      & sId &_
		"&sTitle="   & sTitle &_
		"&sContant=" & sContant &_
		"&sWord="    & sWord


Call Expires()
Call dbopen()

	if (alertMsg <> "")	then
		actType	= ""
	Elseif (actType = "INSERT") Then	'글작성
		Call Insert()
		alertMsg = "입력 되었습니다."
	
	ElseIf (actType = "MODIFY") Then	'글수정
		Call Update()
		If FI_CNT_PWD > 0 Then 
			alertMsg = "수정 되었습니다."
		Else
			Call msgbox("비밀번호가 틀립니다. 비밀번호를 확인해주세요.",true)
		End If
	ElseIf (actType = "DELETE") Then	'글삭제
		Call Delete()
		If FI_CNT_PWD > 0 Then 
			alertMsg = "삭제 되었습니다."
		Else
			Call msgbox("비밀번호가 틀립니다. 비밀번호를 확인해주세요.",true)
		End If		
	else
		alertMsg = "actType[" & actType & "]이 정의되지 않았습니다."
	end If
	
Call dbclose()

'입력
Sub Insert()
	SET objCmd	= Server.CreateObject("ADODB.Command")

SQL = "SET NOCOUNT ON; " &_
	"INSERT INTO [dbo].[SP_JOB_COMPANY]( " &_
	"	 [OwnerName] " &_
	"	,[ManagerName] " &_
	"	,[HomePage] " &_
	"	,[CompanyName] " &_
	"	,[Addr] " &_
	"	,[Tel] " &_
	"	,[Fax] " &_
	"	,[Email] " &_
	"	,[Title] " &_
	"	,[Form] " &_
	"	,[Kind] " &_
	"	,[WorkArea] " &_
	"	,[WorkTime] " &_
	"	,[StaffCnt] " &_
	"	,[Qualify] " &_
	"	,[Files] " &_
	"	,[Dates] " &_
	"	,[Method] " &_
	"	,[Pay] " &_
	"	,[insure] " &_
	"	,[Bigo] " &_
	"	,[InData] " &_
	"	,[UserIdx] " &_
	"	,[Pwd] " &_
	"	,[Ip] " &_
	"	,[Dellfg] " &_
	")VALUES(" &_
	"	 ? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,getDate() " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,0 " &_
	");"
	
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@OwnerName"   ,adVarChar     , adParamInput, 50         , OwnerName )
		.Parameters.Append .CreateParameter( "@ManagerName" ,adVarChar     , adParamInput, 50         , ManagerName )
		.Parameters.Append .CreateParameter( "@HomePage"    ,adVarChar     , adParamInput, 200        , HomePage )
		.Parameters.Append .CreateParameter( "@CompanyName" ,adVarChar     , adParamInput, 200        , CompanyName )
		.Parameters.Append .CreateParameter( "@Addr"        ,adVarChar     , adParamInput, 200        , Addr )
		.Parameters.Append .CreateParameter( "@Tel"         ,adVarChar     , adParamInput, 100        , Tel )
		.Parameters.Append .CreateParameter( "@Fax"         ,adVarChar     , adParamInput, 100        , Fax )
		.Parameters.Append .CreateParameter( "@Email"       ,adVarChar     , adParamInput, 200        , Email )
		.Parameters.Append .CreateParameter( "@Title"       ,adVarChar     , adParamInput, 200        , Title )
		.Parameters.Append .CreateParameter( "@Form"        ,adInteger     , adParamInput, 0          , Form )
		.Parameters.Append .CreateParameter( "@Kind"        ,adInteger     , adParamInput, 0          , Kind )
		.Parameters.Append .CreateParameter( "@WorkArea"    ,adVarChar     , adParamInput, 200        , WorkArea )
		.Parameters.Append .CreateParameter( "@WorkTime"    ,adVarChar     , adParamInput, 200        , WorkTime )
		.Parameters.Append .CreateParameter( "@StaffCnt"    ,adInteger     , adParamInput, 0          , StaffCnt )
		.Parameters.Append .CreateParameter( "@Qualify"     ,adVarChar     , adParamInput, 200        , Qualify )
		.Parameters.Append .CreateParameter( "@Files"       ,adVarChar     , adParamInput, 200        , Files )
		.Parameters.Append .CreateParameter( "@Dates"       ,adVarChar     , adParamInput, 20         , Dates )
		.Parameters.Append .CreateParameter( "@Method"      ,adVarChar     , adParamInput, 200        , Method )
		.Parameters.Append .CreateParameter( "@Pay"         ,adVarChar     , adParamInput, 200        , Pay )
		.Parameters.Append .CreateParameter( "@insure"      ,adInteger     , adParamInput, 0          , insure )
		.Parameters.Append .CreateParameter( "@Bigo"        ,adLongVarChar , adParamInput, 2147483647 , Bigo )

		.Parameters.Append .CreateParameter( "@UserIdx"     ,adInteger     , adParamInput, 0          , UserIdx )
		.Parameters.Append .CreateParameter( "@Pwd"         ,adVarChar     , adParamInput, 50         , Pwd )
		.Parameters.Append .CreateParameter( "@g_uip"       ,adVarChar     , adParamInput, 20         , g_uip )
		.Execute
	End with
	call cmdclose()
End Sub

'수정
Sub Update()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_
	
	"DECLARE @CNT_PWD INT , @PWD VARCHAR(50) , @IDX INT;" &_
	"SET @IDX = ?; " &_
	"SET @PWD = ?; " &_
	"SET @CNT_PWD = (SELECT COUNT(*) FROM [dbo].[SP_JOB_COMPANY] WHERE [Idx] = @IDX AND [Pwd] = @PWD); " &_
	
	"IF @CNT_PWD > 0 " &_
	"BEGIN " &_
	"	UPDATE [dbo].[SP_JOB_COMPANY] SET " &_
	"		 [OwnerName]   = ? " &_
	"		,[ManagerName] = ? " &_
	"		,[HomePage]    = ? " &_
	"		,[CompanyName] = ? " &_
	"		,[Addr]        = ? " &_
	"		,[Tel]         = ? " &_
	"		,[Fax]         = ? " &_
	"		,[Email]       = ? " &_
	"		,[Title]       = ? " &_
	"		,[Form]        = ? " &_
	"		,[Kind]        = ? " &_
	"		,[WorkArea]    = ? " &_
	"		,[WorkTime]    = ? " &_
	"		,[StaffCnt]    = ? " &_
	"		,[Qualify]     = ? " &_
	"		,[Files]       = ? " &_
	"		,[Dates]       = ? " &_
	"		,[Method]      = ? " &_
	"		,[Pay]         = ? " &_
	"		,[insure]      = ? " &_
	"		,[Bigo]        = ? " &_
	"		,[UserIdx]     = ? " &_
	"		,[Ip]          = ? " &_
	"	WHERE [Idx] = @IDX " &_
	"END; " &_
	"SELECT @CNT_PWD AS [CNT_PWD] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"         ,adInteger     , adParamInput, 0          , Idx )
		.Parameters.Append .CreateParameter( "@Pwd"         ,adVarChar     , adParamInput, 50         , Pwd )

		.Parameters.Append .CreateParameter( "@OwnerName"   ,adVarChar     , adParamInput, 50         , OwnerName )
		.Parameters.Append .CreateParameter( "@ManagerName" ,adVarChar     , adParamInput, 50         , ManagerName )
		.Parameters.Append .CreateParameter( "@HomePage"    ,adVarChar     , adParamInput, 200        , HomePage )
		.Parameters.Append .CreateParameter( "@CompanyName" ,adVarChar     , adParamInput, 200        , CompanyName )
		.Parameters.Append .CreateParameter( "@Addr"        ,adVarChar     , adParamInput, 200        , Addr )
		.Parameters.Append .CreateParameter( "@Tel"         ,adVarChar     , adParamInput, 100        , Tel )
		.Parameters.Append .CreateParameter( "@Fax"         ,adVarChar     , adParamInput, 100        , Fax )
		.Parameters.Append .CreateParameter( "@Email"       ,adVarChar     , adParamInput, 200        , Email )
		.Parameters.Append .CreateParameter( "@Title"       ,adVarChar     , adParamInput, 200        , Title )
		.Parameters.Append .CreateParameter( "@Form"        ,adInteger     , adParamInput, 0          , Form )
		.Parameters.Append .CreateParameter( "@Kind"        ,adInteger     , adParamInput, 0          , Kind )
		.Parameters.Append .CreateParameter( "@WorkArea"    ,adVarChar     , adParamInput, 200        , WorkArea )
		.Parameters.Append .CreateParameter( "@WorkTime"    ,adVarChar     , adParamInput, 200        , WorkTime )
		.Parameters.Append .CreateParameter( "@StaffCnt"    ,adInteger     , adParamInput, 0          , StaffCnt )
		.Parameters.Append .CreateParameter( "@Qualify"     ,adVarChar     , adParamInput, 200        , Qualify )
		.Parameters.Append .CreateParameter( "@Files"       ,adVarChar     , adParamInput, 200        , Files )
		.Parameters.Append .CreateParameter( "@Dates"       ,adVarChar     , adParamInput, 20         , Dates )
		.Parameters.Append .CreateParameter( "@Method"      ,adVarChar     , adParamInput, 200        , Method )
		.Parameters.Append .CreateParameter( "@Pay"         ,adVarChar     , adParamInput, 200        , Pay )
		.Parameters.Append .CreateParameter( "@insure"      ,adInteger     , adParamInput, 0          , insure )
		.Parameters.Append .CreateParameter( "@Bigo"        ,adLongVarChar , adParamInput, 2147483647 , Bigo )
		.Parameters.Append .CreateParameter( "@UserIdx"     ,adInteger     , adParamInput, 0          , UserIdx )		
		.Parameters.Append .CreateParameter( "@g_uip"       ,adVarChar     , adParamInput, 20         , g_uip )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
'삭제
Sub Delete()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	
	"DECLARE @CNT_PWD INT , @PWD VARCHAR(50) , @IDX INT;" &_
	"SET @IDX = ?; " &_
	"SET @PWD = ?; " &_
	"SET @CNT_PWD = (SELECT COUNT(*) FROM [dbo].[SP_JOB_COMPANY] WHERE [Idx] = @IDX AND [Pwd] = @PWD); " &_
	
	"IF @CNT_PWD > 0 " &_
	"BEGIN " &_
	"	UPDATE [dbo].[SP_JOB_COMPANY] SET " &_
	"		[Dellfg] = 1 " &_
	"	WHERE [Idx] = @IDX " &_
	"END; " &_

	"SELECT @CNT_PWD AS [CNT_PWD] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput, 0  , Idx )
		.Parameters.Append .CreateParameter( "@Pwd" ,adVarChar , adParamInput, 50 , Pwd )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN">
<html>
<head>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
</head>
<script language=javascript>
	if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
	top.location.href = "fStaffL.asp?<%=PageParams%>";
</script>
</html>