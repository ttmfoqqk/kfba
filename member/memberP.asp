<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim CoUrl       : CoUrl        = "/"
Dim alertMsg    : alertMsg     = ""
Dim actType     : actType      = RequestSet("actType"     , "POST" , "")
Dim UserId      : UserId       = RequestSet("UserId"      , "POST" , "")
Dim UserPwd     : UserPwd      = RequestSet("UserPwd"     , "POST" , "")
Dim UserPwdc    : UserPwdc     = RequestSet("UserPwdc"    , "POST" , "")
Dim UserName    : UserName     = RequestSet("UserName"    , "POST" , "")
Dim UserBirth1  : UserBirth1   = RequestSet("UserBirth1"  , "POST" , "")
Dim UserBirth2  : UserBirth2   = RequestSet("UserBirth2"  , "POST" , "")
Dim UserBirth3  : UserBirth3   = RequestSet("UserBirth3"  , "POST" , "")
Dim UserPhone1  : UserPhone1   = RequestSet("UserPhone1"  , "POST" , "")
Dim UserPhone2  : UserPhone2   = RequestSet("UserPhone2"  , "POST" , "")
Dim UserPhone3  : UserPhone3   = RequestSet("UserPhone3"  , "POST" , "")
Dim UserSmsFg   : UserSmsFg    = RequestSet("UserSmsFg"   , "POST" , 0 )
Dim UserEmail1  : UserEmail1   = RequestSet("UserEmail1"  , "POST" , "")
Dim UserEmail3  : UserEmail3   = RequestSet("UserEmail3"  , "POST" , "")
Dim UserEmailfg : UserEmailfg  = RequestSet("UserEmailfg" , "POST" , 0 )
Dim UserZip1    : UserZip1     = RequestSet("UserZip1"    , "POST" , "")
Dim UserZip2    : UserZip2     = RequestSet("UserZip2"    , "POST" , "")
Dim UserAddr1   : UserAddr1    = RequestSet("UserAddr1"   , "POST" , "")
Dim UserAddr2   : UserAddr2    = RequestSet("UserAddr2"   , "POST" , "")

Dim LastName    : LastName     = RequestSet("LastName"    , "POST" , "")
Dim FirstName   : FirstName    = RequestSet("FirstName"   , "POST" , "")


Call Expires()
Call dbopen()
	If actType = "INSERT" Then 
		
		If UserPwd <> UserPwdc Then 
			Call msgbox("비밀번호가 잘못입력되었습니다.",true)
		End If

		' // 본인인증 세션 검증

		If session("sName") = "" Or session("sDupInfo") = "" Then 
			Call msgbox("인증정보가 잘못입력되었습니다.",true)
		End If


		Call Insert()
		If FI_IN_CNT > 0 Then 
			Call msgbox("중복된 아이디는 사용하실수 없습니다.",true)
		End If
		CoUrl = "joinOk.asp"

		session("sVNumber")      = ""
		session("sName")         = ""
		session("sBirthDate")    = ""
		session("sGender")       = ""
		session("sNationalInfo") = ""
		session("sDupInfo")      = ""
		session("sConnInfo")     = ""

		'Call sendSmsEmail( UserId, UserEmail1 &"@"& UserEmail3 , "" )
	ElseIf actType = "UPDATE" Then 
		
		If Isnull(Session("UserIdx")) Or Session("UserIdx") = "" Then 
			Call msgbox("로그인 세션이 만료되었습니다",true)
		End If

		Call Update()
		alertMsg = "수정되었습니다."
		CoUrl    = "../mypage/info.asp"
	ElseIf actType = "DELETE" Then 

		If Isnull(Session("UserIdx")) Or Session("UserIdx") = "" Then 
			Call msgbox("로그인 세션이 만료되었습니다",true)
		End If

		Call Delete()

		If FI_CHECK_PASS = 1 Then 
			Session("UserIdx")	= ""
			Session("UserId")	= ""
			Session.Contents.RemoveAll()
			Session.Abandon()
			alertMsg = "탈퇴처리 되었습니다."
			CoUrl    = FRONT_ROOT_DIR
		Else
			With Response
			 .Write "<script language='javascript' type='text/javascript'>"
			 .Write "alert('비밀번호가 정확하지 않습니다.');"
			 .Write "history.go(-1);"
			 .Write "</script>"
			 .End
			End With
		End If		
	Else
		alertMsg = "ERR : [actType] 이 없습니다."
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

	"		,[UserSex] " &_
	"		,[UserDIKEY] " &_
	"		,[UserVssn] " &_
	"		,[UserNational] " &_

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
	"		,? " &_
	"		,? " &_
	"		,? " &_

	"	) " &_
	"END " &_
	"SELECT @IN_CNT AS [IN_CNT] "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserId"      ,adVarChar     , adParamInput, 50, UserId )

		.Parameters.Append .CreateParameter( "@UserId"      ,adVarChar     , adParamInput, 50         , UserId )
		.Parameters.Append .CreateParameter( "@UserPass"    ,adVarChar     , adParamInput, 50         , UserPwd )
		.Parameters.Append .CreateParameter( "@UserName"    ,adVarChar     , adParamInput, 50         , session("sName") )
		.Parameters.Append .CreateParameter( "@UserBirth"   ,adChar        , adParamInput, 8          , session("sBirthDate") )
		.Parameters.Append .CreateParameter( "@UserHphone1" ,adChar        , adParamInput, 4          , UserPhone1 )
		.Parameters.Append .CreateParameter( "@UserHphone2" ,adChar        , adParamInput, 4          , UserPhone2 )
		.Parameters.Append .CreateParameter( "@UserHphone3" ,adChar        , adParamInput, 4          , UserPhone3 )
		.Parameters.Append .CreateParameter( "@UserSmsFg"   ,adInteger     , adParamInput, 0          , UserSmsFg )
		.Parameters.Append .CreateParameter( "@UserEmail"   ,adVarChar     , adParamInput, 200        , UserEmail1 & "@" & UserEmail3 )
		.Parameters.Append .CreateParameter( "@UserEmailFg" ,adInteger     , adParamInput, 0          , UserEmailfg )
		.Parameters.Append .CreateParameter( "@UserZipcode" ,adChar        , adParamInput, 6          , UserZip1 & UserZip2 )
		.Parameters.Append .CreateParameter( "@UserAddr1"   ,adVarChar     , adParamInput, 200        , UserAddr1 )
		.Parameters.Append .CreateParameter( "@UserAddr2"   ,adVarChar     , adParamInput, 200        , UserAddr2 )

		.Parameters.Append .CreateParameter( "@UserSex"      ,adInteger     , adParamInput, 0         , session("sGender") )
		.Parameters.Append .CreateParameter( "@UserDIKEY"    ,adVarChar     , adParamInput, 64        , session("sDupInfo") )
		.Parameters.Append .CreateParameter( "@UserVssn"     ,adVarChar     , adParamInput, 13        , IIF( session("sVNumber")="","",session("sVNumber") ) )
		.Parameters.Append .CreateParameter( "@UserNational" ,adInteger     , adParamInput, 0         , session("sNationalInfo") )

		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "UPDATE [dbo].[SP_USER_MEMBER] SET " &_
	"	 [UserHphone1] = ? " &_
	"	,[UserHphone2] = ? " &_
	"	,[UserHphone3] = ? " &_
	"	,[UserSmsFg]   = ? " &_
	"	,[UserEmail]   = ? " &_
	"	,[UserEmailFg] = ? " &_
	"	,[UserZipcode] = ? " &_
	"	,[UserAddr1]   = ? " &_
	"	,[UserAddr2]   = ? " &_
	"	,[FirstName]   = ? " &_
    "	,[LastName]    = ? " &_
	"WHERE [UserIdx]   = ? "
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserHphone1" ,adChar        , adParamInput, 4          , UserPhone1 )
		.Parameters.Append .CreateParameter( "@UserHphone2" ,adChar        , adParamInput, 4          , UserPhone2 )
		.Parameters.Append .CreateParameter( "@UserHphone3" ,adChar        , adParamInput, 4          , UserPhone3 )
		.Parameters.Append .CreateParameter( "@UserSmsFg"   ,adInteger     , adParamInput, 0          , UserSmsFg )
		.Parameters.Append .CreateParameter( "@UserEmail"   ,adVarChar     , adParamInput, 200        , UserEmail1 & "@" & UserEmail3 )
		.Parameters.Append .CreateParameter( "@UserEmailFg" ,adInteger     , adParamInput, 0          , UserEmailfg )
		.Parameters.Append .CreateParameter( "@UserZipcode" ,adChar        , adParamInput, 6          , UserZip1 & UserZip2 )
		.Parameters.Append .CreateParameter( "@UserAddr1"   ,adVarChar     , adParamInput, 200        , UserAddr1 )
		.Parameters.Append .CreateParameter( "@UserAddr2"   ,adVarChar     , adParamInput, 200        , UserAddr2 )
		.Parameters.Append .CreateParameter( "@FirstName"   ,adVarChar     , adParamInput, 50         , FirstName )
		.Parameters.Append .CreateParameter( "@LastName"    ,adVarChar     , adParamInput, 50         , LastName )
		.Parameters.Append .CreateParameter( "@UserIdx"     ,adInteger     , adParamInput, 0          , Session("UserIdx") )
		.Execute
	End with
	call cmdclose()
End Sub

'탈퇴
Sub Delete()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	
	SQL = "set nocount on; " &_
	"DECLARE @CHECK_PASS INT " &_
	"DECLARE @UserIdx INT , @UserPass VARCHAR(50) ;" &_
	"SET @UserIdx = ? " &_
	"SET @UserPass = ? " &_

	"SELECT " &_
	"	 @CHECK_PASS	= pwdcompare(@UserPass,[UserPass]) " &_
	"FROM [dbo].[SP_USER_MEMBER] " &_
	"WHERE [UserIdx] = @UserIdx " &_

	"IF @CHECK_PASS = 1 " &_
	"BEGIN " &_
	"	UPDATE [dbo].[SP_USER_MEMBER] SET " &_
	"		 [UserDelfg] = 1 " &_
	"		,[UserDelfgDate] = getdate() " &_
	"	where UserIdx = @UserIdx " &_
	"END " &_
	
	"SELECT @CHECK_PASS AS [CHECK_PASS] " 

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx"  ,adInteger , adParamInput  , 0 , Session("UserIdx") )
		.Parameters.Append .CreateParameter( "@UserPass" ,adVarChar , adParamInput  , 50 , UserPwd )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub


Sub sendSmsEmail(UserId, user_email, attachPath)

	Dim strFile     : strFile     = server.mapPath(FRONT_ROOT_DIR & "_skin/mail/" ) & "/mail_join.html"
	Dim strTitle    : strTitle    = SEND_MAIL_NAME & " " & UserId & "님 회원가입이 정상적으로 처리 되었습니다."
	Dim strContants : strContants = SEND_MAIL_NAME & " " & UserId & "님 회원가입이 정상적으로 처리 되었습니다."

	Dim mfrom		: mfrom		= SITE_NAME & " " & SEND_MAIL_MAIL
	Dim mto			: mto		= user_email
	Dim mtitle		: mtitle	= strTitle
	Dim mcontents	: mcontents	= ReadFile(strFile)
		mcontents	= replace(mcontents, "#CONTANTS#", strContants )
		mcontents	= replace(mcontents, "#BOTTOM_INFO#", SEND_MAIL_BOTTOM_INFO )
		mcontents	= replace(mcontents, "#BOTTOM_COPY#", SEND_MAIL_BOTTOM_COPY )
	Dim mailMessage : mailMessage = MailSend(mtitle, mcontents, mto, mfrom, attachPath)
	
	if LEN(mailMessage) > 0 then
		'response.write "메일 전송 에러 : " & mailMessage
		'response.end
	end if
End Sub	

%>
<!DOCTYPE html>
<HTML>
<HEAD>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
</HEAD>
<script language=javascript>
	if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
	top.location.href = "<%=CoUrl%>";
</script>
</html>