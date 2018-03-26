<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim alertMsg     : alertMsg     = ""
Dim adminIdx     : adminIdx     = Request.Form("adminIdx")
Dim actType      : actType      = Trim( Request.Form("actType") )
Dim adminId      : adminId      = Trim( Request.Form("adminId") )
Dim adminPwd     : adminPwd     = Trim( Request.Form("adminPwd") )
Dim adminPwdCk   : adminPwdCk   = Trim( Request.Form("adminPwdCk") )
Dim adminName    : adminName    = Trim( Request.Form("adminName") )
Dim adminPhone1  : adminPhone1  = Trim( Request.Form("adminPhone1") )
Dim adminPhone2  : adminPhone2  = Trim( Request.Form("adminPhone2") )
Dim adminPhone3  : adminPhone3  = Trim( Request.Form("adminPhone3") )
Dim adminHphone1 : adminHphone1 = Trim( Request.Form("adminHphone1") )
Dim adminHphone2 : adminHphone2 = Trim( Request.Form("adminHphone2") )
Dim adminHphone3 : adminHphone3 = Trim( Request.Form("adminHphone3") )
Dim adminExt     : adminExt     = Trim( Request.Form("adminExt") )
Dim adminDir     : adminDir     = Trim( Request.Form("adminDir") )
Dim adminMail1   : adminMail1   = Trim( Request.Form("adminMail1") )
Dim adminMail3   : adminMail3   = Trim( Request.Form("adminMail3") )
Dim adminMail    : adminMail    = adminMail1 & "@" & adminMail3
Dim adminMsg1    : adminMsg1    = Trim( Request.Form("adminMsg1") )
Dim adminMsg3    : adminMsg3    = Trim( Request.Form("adminMsg3") )
Dim adminMsg     : adminMsg     = adminMsg1 & "@" & adminMsg3
Dim adminBigo    : adminBigo    = Trim( TagEncode(Request.Form("adminBigo")) )

Dim pageNo        : pageNo      = Request.Form("pageNo")



If adminPwd <> adminPwdCk Then 
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('비밀번호를 확인해주세요.');"
	 .Write "history.back(-1);"
	 .Write "</script>"
	 .End
	End With
End If

Call Expires()
Call dbopen()
	If actType = "INSERT" Then 
		Call Insert()
		If FI_IN_CNT > 0 Then 
			alertMsg = "중복된 아이디는 사용하실수 없습니다."
		Else
			alertMsg = "입력되었습니다."
		End If
	ElseIf actType = "UPDATE" Then 
		Call Update()
		alertMsg = "수정되었습니다."
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

	SQL = "set nocount on;DECLARE @IN_CNT INT;" &_
	"SET @IN_CNT = ( " &_
	"	SELECT COUNT(*) FROM [dbo].[SP_ADMIN_MEMBER] WHERE [Id] = ? " &_
	") " &_
	"IF @IN_CNT = 0 " &_
	"	BEGIN " &_
	
	"	INSERT INTO [dbo].[SP_ADMIN_MEMBER]( " &_
	"		 [Id] " &_
	"		,[Pwd] " &_
	"		,[Name] " &_
	"		,[pHone1] " &_
	"		,[pHone2] " &_
	"		,[pHone3] " &_
	"		,[Hphone1] " &_
	"		,[Hphone2] " &_
	"		,[Hphone3] " &_
	"		,[ExtNum] " &_
	"		,[DirNum] " &_
	"		,[email] " &_
	"		,[MsgAddr] " &_
	"		,[Bigo] " &_
	"		,[Indata] " &_
	"		,[Dellfg] " &_
	"	)VALUES( " &_
	"		 ? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,? " &_
	"		,getdate() " &_
	"		,0 " &_
	"	) " &_
	"END " &_
	"SELECT @IN_CNT AS [IN_CNT] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Id"      ,adVarChar     , adParamInput, 50, adminId )

		.Parameters.Append .CreateParameter( "@Id"      ,adVarChar     , adParamInput, 50         , adminId )
		.Parameters.Append .CreateParameter( "@Pwd"     ,adVarChar     , adParamInput, 50         , adminPwd )
		.Parameters.Append .CreateParameter( "@Name"    ,adVarChar     , adParamInput, 50         , adminName )
		.Parameters.Append .CreateParameter( "@pHone1"  ,adChar        , adParamInput, 4          , adminPhone1 )
		.Parameters.Append .CreateParameter( "@pHone2"  ,adChar        , adParamInput, 4          , adminPhone2 )
		.Parameters.Append .CreateParameter( "@pHone3"  ,adChar        , adParamInput, 4          , adminPhone3 )
		.Parameters.Append .CreateParameter( "@Hphone1" ,adChar        , adParamInput, 4          , adminHphone1 )
		.Parameters.Append .CreateParameter( "@Hphone2" ,adChar        , adParamInput, 4          , adminHphone2 )
		.Parameters.Append .CreateParameter( "@Hphone3" ,adChar        , adParamInput, 4          , adminHphone3 )
		.Parameters.Append .CreateParameter( "@ExtNum"  ,adVarChar     , adParamInput, 50         , adminExt )
		.Parameters.Append .CreateParameter( "@DirNum"  ,adVarChar     , adParamInput, 50         , adminDir )
		.Parameters.Append .CreateParameter( "@email"   ,adVarChar     , adParamInput, 200        , adminMail )
		.Parameters.Append .CreateParameter( "@MsgAddr" ,adVarChar     , adParamInput, 200        , adminMsg )
		.Parameters.Append .CreateParameter( "@Bigo"    ,adLongVarChar , adParamInput, 2147483647 , adminBigo )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "UPDATE [dbo].[SP_ADMIN_MEMBER] SET " &_
	"	 [Pwd]     = ? " &_
	"	,[Name]    = ? " &_
	"	,[pHone1]  = ? " &_
	"	,[pHone2]  = ? " &_
	"	,[pHone3]  = ? " &_
	"	,[Hphone1] = ? " &_
	"	,[Hphone2] = ? " &_
	"	,[Hphone3] = ? " &_
	"	,[ExtNum]  = ? " &_
	"	,[DirNum]  = ? " &_
	"	,[email]   = ? " &_
	"	,[MsgAddr] = ? " &_
	"	,[Bigo]    = ? " &_
	"WHERE [Idx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Pwd"     ,adVarChar     , adParamInput, 50         , adminPwd )
		.Parameters.Append .CreateParameter( "@Name"    ,adVarChar     , adParamInput, 50         , adminName )
		.Parameters.Append .CreateParameter( "@pHone1"  ,adChar        , adParamInput, 4          , adminPhone1 )
		.Parameters.Append .CreateParameter( "@pHone2"  ,adChar        , adParamInput, 4          , adminPhone2 )
		.Parameters.Append .CreateParameter( "@pHone3"  ,adChar        , adParamInput, 4          , adminPhone3 )
		.Parameters.Append .CreateParameter( "@Hphone1" ,adChar        , adParamInput, 4          , adminHphone1 )
		.Parameters.Append .CreateParameter( "@Hphone2" ,adChar        , adParamInput, 4          , adminHphone2 )
		.Parameters.Append .CreateParameter( "@Hphone3" ,adChar        , adParamInput, 4          , adminHphone3 )
		.Parameters.Append .CreateParameter( "@ExtNum"  ,adVarChar     , adParamInput, 50         , adminExt )
		.Parameters.Append .CreateParameter( "@DirNum"  ,adVarChar     , adParamInput, 50         , adminDir )
		.Parameters.Append .CreateParameter( "@email"   ,adVarChar     , adParamInput, 200        , adminMail )
		.Parameters.Append .CreateParameter( "@MsgAddr" ,adVarChar     , adParamInput, 200        , adminMsg )
		.Parameters.Append .CreateParameter( "@Bigo"    ,adLongVarChar , adParamInput, 2147483647 , adminBigo )
		.Parameters.Append .CreateParameter( "@Idx"     ,adInteger     , adParamInput, 0          , adminIdx )
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
	
	
	"UPDATE [dbo].[SP_ADMIN_MEMBER] SET " &_
	"	[Dellfg] = 1 " &_
	"WHERE [Idx] in( SELECT T_INT FROM @T ) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adVarChar , adParamInput, 8000 , adminIdx )
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
	top.location.href = "companyMemberL.asp?pageNo=<%=pageNo%>";
</script>
</html>