<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim alertMsg     : alertMsg     = ""
Dim UserIdx      : UserIdx      = RequestSet("UserIdx" ,"POST", "")
Dim actType      : actType      = RequestSet("actType" ,"POST", "")
Dim UserId       : UserId       = RequestSet("UserId" ,"POST", "")
Dim UserPass     : UserPass     = RequestSet("UserPass" ,"POST", "")
Dim UserName     : UserName     = RequestSet("UserName" ,"POST", "")
Dim UserBirth1   : UserBirth1   = RequestSet("UserBirth1" ,"POST", "")
Dim UserBirth2   : UserBirth2   = RequestSet("UserBirth2" ,"POST", "")
Dim UserBirth3   : UserBirth3   = RequestSet("UserBirth3" ,"POST", "")
Dim UserHphone1  : UserHphone1  = RequestSet("UserHphone1" ,"POST", "")
Dim UserHphone2  : UserHphone2  = RequestSet("UserHphone2" ,"POST", "")
Dim UserHphone3  : UserHphone3  = RequestSet("UserHphone3" ,"POST", "")
Dim UserSmsFg    : UserSmsFg    = RequestSet("UserSmsFg" ,"POST", 0)
Dim UserEmail1   : UserEmail1   = RequestSet("UserEmail1" ,"POST", "")
Dim UserEmail3   : UserEmail3   = RequestSet("UserEmail3" ,"POST", "")
Dim UserEmailFg  : UserEmailFg  = RequestSet("UserEmailFg" ,"POST", 0)
Dim UserZipcode1 : UserZipcode1 = RequestSet("UserZipcode1" ,"POST", "")
Dim UserZipcode2 : UserZipcode2 = RequestSet("UserZipcode2" ,"POST", "")
Dim UserAddr1    : UserAddr1    = RequestSet("UserAddr1" ,"POST", "")
Dim UserAddr2    : UserAddr2    = RequestSet("UserAddr2" ,"POST", "")
Dim UserBigo     : UserBigo     = TagEncode( RequestSet("UserBigo","POST","") )
Dim UserDelFg    : UserDelFg    = RequestSet("UserDelFg" ,"POST", 0)

Dim pageNo     : pageNo     = RequestSet("pageNo","POST",1)
Dim sIndate    : sIndate    = RequestSet("sIndate","POST","")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate","POST","")
Dim sUserId    : sUserId    = RequestSet("sUserId","POST","")
Dim sUserName  : sUserName  = RequestSet("sUserName","POST","")
Dim sHphone3   : sHphone3   = RequestSet("sHphone3","POST","")
Dim sUserBirth : sUserBirth = RequestSet("sUserBirth","POST","")
Dim sState     : sState     = RequestSet("sState","POST","")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sUserId="    & sUserId &_
		"&sUserName="  & sUserName &_
		"&sHphone3="   & sHphone3 &_
		"&sUserBirth=" & sUserBirth &_
		"&sState="     & sState


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
		If FI_IN_CNT > 0 Then 
			alertMsg = "중복된 아이디는 사용하실수 없습니다."
		Else
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

	SQL = "set nocount on;DECLARE @IN_CNT INT;" &_
	"SET @IN_CNT = ( " &_
	"	SELECT COUNT(*) FROM [dbo].[SP_USER_MEMBER] WHERE [UserId] = ? " &_
	") " &_
	"IF @IN_CNT = 0 " &_
	"	BEGIN " &_
	"	INSERT INTO [dbo].[SP_USER_MEMBER]( " &_
	"		 [UserId] " &_
	"		,[UserPass] " &_
	"		,[UserName] " &_
	"		,[UserBirth] " &_
	"		,[UserHphone1] " &_
	"		,[UserHphone2] " &_
	"		,[UserHphone3] " &_
	"		,[UserSmsFg] " &_
	"		,[UserEmail] " &_
	"		,[UserEmailFg] " &_
	"		,[UserZipcode] " &_
	"		,[UserAddr1] " &_
	"		,[UserAddr2] " &_
	"		,[UserIndate] " &_
	"		,[UserDelFg] " &_
	"		,[UserBigo] " &_
	"	)VALUES( " &_
	"		 ? " &_
	"		,pwdencrypt(?) " &_
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
	"		,? " &_
	"	) " &_
	"END " &_
	"SELECT @IN_CNT AS [IN_CNT] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserId"      ,adVarChar     , adParamInput, 50, UserId )

		.Parameters.Append .CreateParameter( "@UserId"      ,adVarChar     , adParamInput, 50         , UserId )
		.Parameters.Append .CreateParameter( "@UserPass"    ,adVarChar     , adParamInput, 50         , UserPass )
		.Parameters.Append .CreateParameter( "@UserName"    ,adVarChar     , adParamInput, 50         , UserName )
		.Parameters.Append .CreateParameter( "@UserBirth"   ,adChar        , adParamInput, 8          , UserBirth1 & UserBirth2 & UserBirth3 )
		.Parameters.Append .CreateParameter( "@UserHphone1" ,adChar        , adParamInput, 4          , UserHphone1 )
		.Parameters.Append .CreateParameter( "@UserHphone2" ,adChar        , adParamInput, 4          , UserHphone2 )
		.Parameters.Append .CreateParameter( "@UserHphone3" ,adChar        , adParamInput, 4          , UserHphone3 )
		.Parameters.Append .CreateParameter( "@UserSmsFg"   ,adInteger     , adParamInput, 0          , UserSmsFg )
		.Parameters.Append .CreateParameter( "@UserEmail"   ,adVarChar     , adParamInput, 200        , UserEmail1 & "@" & UserEmail3 )
		.Parameters.Append .CreateParameter( "@UserEmailFg" ,adInteger     , adParamInput, 0          , UserEmailFg )
		.Parameters.Append .CreateParameter( "@UserZipcode" ,adChar        , adParamInput, 6          , UserZipcode1 & UserZipcode2 )
		.Parameters.Append .CreateParameter( "@UserAddr1"   ,adVarChar     , adParamInput, 200        , UserAddr1 )
		.Parameters.Append .CreateParameter( "@UserAddr2"   ,adVarChar     , adParamInput, 200        , UserAddr2 )
		.Parameters.Append .CreateParameter( "@UserBigo"    ,adLongVarChar , adParamInput, 2147483647 , UserBigo )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
'수정
Sub Update()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "set nocount on;DECLARE @IN_CNT INT,@UserId varchar(50),@UserIdx INT;" &_
	"SET @UserId = ? " &_
	"SET @UserIdx = ? " &_
	"SET @IN_CNT = ( " &_
	"	SELECT COUNT(*) FROM [dbo].[SP_USER_MEMBER] WHERE [UserId] = @UserId AND [UserIdx] != @UserIdx " &_
	") " &_
	"IF @IN_CNT = 0 " &_
	"	BEGIN " &_
	"	UPDATE [dbo].[SP_USER_MEMBER] SET " &_
	"		 [UserId]      = @UserId " &_
	"		,[UserName]    = ? " &_
	"		,[UserBirth]   = ? " &_
	"		,[UserHphone1] = ? " &_
	"		,[UserHphone2] = ? " &_
	"		,[UserHphone3] = ? " &_
	"		,[UserSmsFg]   = ? " &_
	"		,[UserEmail]   = ? " &_
	"		,[UserEmailFg] = ? " &_
	"		,[UserZipcode] = ? " &_
	"		,[UserAddr1]   = ? " &_
	"		,[UserAddr2]   = ? " &_
	"		,[UserDelFg]   = ? " &_
	"		,[UserBigo]    = ? " &_
	"	WHERE [UserIdx]   = @UserIdx " &_

	"	DECLARE @NEW_PWD VARCHAR(50); " &_
	"	SET @NEW_PWD = ? " &_
	"	IF @NEW_PWD != '' " &_
	"	BEGIN " &_
	"		UPDATE [dbo].[SP_USER_MEMBER] SET " &_
	"			[UserPass]    = pwdencrypt(@NEW_PWD) " &_
	"		WHERE [UserIdx]   = @UserIdx " &_
	"	END " &_

	"END " &_
	"SELECT @IN_CNT AS [IN_CNT] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserId"      ,adVarChar     , adParamInput, 50         , UserId )
		.Parameters.Append .CreateParameter( "@UserIdx"     ,adInteger     , adParamInput, 50         , IIF(UserIdx="",0,UserIdx) )

		.Parameters.Append .CreateParameter( "@UserName"    ,adVarChar     , adParamInput, 50         , UserName )
		.Parameters.Append .CreateParameter( "@UserBirth"   ,adChar        , adParamInput, 8          , UserBirth1 & UserBirth2 & UserBirth3 )
		.Parameters.Append .CreateParameter( "@UserHphone1" ,adChar        , adParamInput, 4          , UserHphone1 )
		.Parameters.Append .CreateParameter( "@UserHphone2" ,adChar        , adParamInput, 4          , UserHphone2 )
		.Parameters.Append .CreateParameter( "@UserHphone3" ,adChar        , adParamInput, 4          , UserHphone3 )
		.Parameters.Append .CreateParameter( "@UserSmsFg"   ,adInteger     , adParamInput, 0          , UserSmsFg )
		.Parameters.Append .CreateParameter( "@UserEmail"   ,adVarChar     , adParamInput, 200        , UserEmail1 & "@" & UserEmail3 )
		.Parameters.Append .CreateParameter( "@UserEmailFg" ,adInteger     , adParamInput, 0          , UserEmailFg )
		.Parameters.Append .CreateParameter( "@UserZipcode" ,adChar        , adParamInput, 6          , UserZipcode1 & UserZipcode2 )
		.Parameters.Append .CreateParameter( "@UserAddr1"   ,adVarChar     , adParamInput, 200        , UserAddr1 )
		.Parameters.Append .CreateParameter( "@UserAddr2"   ,adVarChar     , adParamInput, 200        , UserAddr2 )
		.Parameters.Append .CreateParameter( "@UserDelFg"   ,adInteger     , adParamInput, 0          , UserDelFg )
		.Parameters.Append .CreateParameter( "@UserBigo"    ,adLongVarChar , adParamInput, 2147483647 , UserBigo )

		.Parameters.Append .CreateParameter( "@UserPass"    ,adVarChar     , adParamInput, 50         , UserPass )

		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
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
	
	
	"UPDATE [dbo].[SP_USER_MEMBER] SET " &_
	"	[UserDelFg] = 1 " &_
	"WHERE [UserIdx] in( SELECT T_INT FROM @T ) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adVarChar , adParamInput, 8000 , UserIdx )
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
	top.location.href = "memberL.asp?<%=PageParams%>";
</script>
</html>