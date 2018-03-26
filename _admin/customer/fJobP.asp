<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%

Dim alertMsg     : alertMsg = ""
Dim actType      : actType     = RequestSet("actType"     ,"POST","")

Dim Idx          : Idx         = RequestSet("Idx"         ,"POST",0)
Dim OwnerName    : OwnerName   = RequestSet("OwnerName"   ,"POST","")
Dim ManagerName  : ManagerName = RequestSet("ManagerName" ,"POST","")
Dim HomePage     : HomePage    = RequestSet("HomePage"    ,"POST","")
Dim CompanyName  : CompanyName = RequestSet("CompanyName" ,"POST","")
Dim Addr         : Addr        = RequestSet("Addr"        ,"POST","")
Dim Tel          : Tel         = RequestSet("Tel"         ,"POST","")
Dim Fax          : Fax         = RequestSet("Fax"         ,"POST","")
Dim Email        : Email       = RequestSet("Email"       ,"POST","")
Dim Title        : Title       = RequestSet("Title"       ,"POST","")
Dim Form         : Form        = RequestSet("Form"        ,"POST",0)
Dim Kind         : Kind        = RequestSet("Kind"        ,"POST",0)
Dim WorkArea     : WorkArea    = RequestSet("WorkArea"    ,"POST","")
Dim WorkTime     : WorkTime    = RequestSet("WorkTime"    ,"POST","")
Dim StaffCnt     : StaffCnt    = RequestSet("StaffCnt"    ,"POST",0)
Dim Qualify      : Qualify     = RequestSet("Qualify"     ,"POST","")
Dim Files        : Files       = RequestSet("Files"       ,"POST","")
Dim Dates        : Dates       = RequestSet("Dates"       ,"POST","")
Dim Method       : Method      = RequestSet("Method"      ,"POST","")
Dim Pay          : Pay         = RequestSet("Pay"         ,"POST","")
Dim insure       : insure      = RequestSet("insure"      ,"POST",0)
Dim Bigo         : Bigo        = RequestSet("Bigo"        ,"POST","")
Dim UserIdx      : UserIdx     = RequestSet("UserIdx"     ,"POST",0)
Dim Pwd          : Pwd         = RequestSet("Pwd"         ,"POST","")

Dim pageNo       : pageNo       = RequestSet("pageNo","GET",1)
Dim sIndate      : sIndate      = RequestSet("sIndate"     ,"GET","")
Dim sOutdate     : sOutdate     = RequestSet("sOutdate"    ,"GET","")
Dim sId          : sId          = RequestSet("sId"         ,"GET","")
Dim sName        : sName        = RequestSet("sName"       ,"GET","")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="  & sIndate &_
		"&sOutdate=" & sOutdate &_
		"&sId="      & sId &_
		"&sName="    & sName

Call Expires()
Call dbopen()

	if (alertMsg <> "")	then
		actType	= ""
	Elseif (actType = "INSERT") Then	'글작성
		Call Insert()
		alertMsg = "입력 되었습니다."
	
	ElseIf (actType = "MODIFY") Then	'글수정
		Call Update()
		alertMsg = "수정 되었습니다."
	ElseIf (actType = "DELETE") Then	'글삭제
		Call Delete()
		alertMsg = "삭제 되었습니다."
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
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "UPDATE [dbo].[SP_JOB_COMPANY] SET " &_
	"	 [OwnerName]   = ? " &_
	"	,[ManagerName] = ? " &_
	"	,[HomePage]    = ? " &_
	"	,[CompanyName] = ? " &_
	"	,[Addr]        = ? " &_
	"	,[Tel]         = ? " &_
	"	,[Fax]         = ? " &_
	"	,[Email]       = ? " &_
	"	,[Title]       = ? " &_
	"	,[Form]        = ? " &_
	"	,[Kind]        = ? " &_
	"	,[WorkArea]    = ? " &_
	"	,[WorkTime]    = ? " &_
	"	,[StaffCnt]    = ? " &_
	"	,[Qualify]     = ? " &_
	"	,[Files]       = ? " &_
	"	,[Dates]       = ? " &_
	"	,[Method]      = ? " &_
	"	,[Pay]         = ? " &_
	"	,[insure]      = ? " &_
	"	,[Bigo]        = ? " &_
	"	,[InData]      = ? " &_
	"	,[UserIdx]     = ? " &_
	"	,[Pwd]         = ? " &_
	"	,[Ip]          = ? " &_
	"WHERE [Idx]       = ? "

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
		.Parameters.Append .CreateParameter( "@Idx"         ,adInteger     , adParamInput, 0          , Idx )
		.Execute
	End with
	call cmdclose()
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
	
	
	"UPDATE [dbo].[SP_JOB_USER] SET " &_
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
	top.location.href = "fJobL.asp?<%=PageParams%>";
</script>
</html>