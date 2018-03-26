<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
checkLogin( g_host & g_url )
Dim UserPass      : UserPass      = RequestSet("UserPass" , "POST" , "")
Dim NewUserPass   : NewUserPass   = RequestSet("NewUserPass" , "POST" , "")
Dim UserPassCheck : UserPassCheck = RequestSet("UserPassCheck" , "POST" , "")
Dim alertMsg      : alertMsg = "수정되었습니다."

If UserPass="" Or NewUserPass="" Or UserPassCheck="" Then 
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('ERR : 잘못된 경로 입니다.');"
	 .Write "history.go(-1);"
	 .Write "</script>"
	 .End
	End With
End If

If NewUserPass <> UserPassCheck Then 
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('ERR : 새 비밀번호가 잘못입력되었습니다.');"
	 .Write "history.go(-1);"
	 .Write "</script>"
	 .End
	End With
End If


Call Expires()
Call dbopen()
	Call update()
Call dbclose()

If FI_CHECK_PASS = 1 Then 
	'Call sendSmsEmail( FI_USERID , FI_USERMAIL ) ' 메일 발송
Else
	alertMsg = "ERR : 비밀번호가 정확하지 않습니다."
End If


Sub update()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @CHECK_PASS INT , @USERID VARCHAR(50), @USERMAIL VARCHAR(200); " &_
	"DECLARE @UserIdx INT , @UserPass VARCHAR(50), @NewUserPass VARCHAR(50); " &_
	"SET @UserIdx = ? " &_
	"SET @UserPass = ? " &_
	"SET @NewUserPass = ? " &_
	"SELECT  " &_
	"	 @CHECK_PASS	= pwdcompare(@UserPass,[UserPass]) " &_
	"	,@USERID		= [UserId] " &_
	"	,@USERMAIL		= [UserEmail] " &_
	"FROM [dbo].[SP_USER_MEMBER] " &_
	"WHERE [UserIdx] = @UserIdx " &_

	"IF @CHECK_PASS = 1 " &_
	"BEGIN " &_
	"	UPDATE [dbo].[SP_USER_MEMBER] SET " &_
	"		[UserPass] = pwdencrypt(@NewUserPass) " &_
	"	where [UserIdx] = @UserIdx " &_
	"END " &_

	"SELECT  " &_
	"	 @CHECK_PASS AS [CHECK_PASS] " &_
	"	,@USERID AS [USERID] " &_
	"	,@USERMAIL AS [USERMAIL] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx"     ,adInteger , adParamInput , 0  , Session("UserIdx") )
		.Parameters.Append .CreateParameter( "@UserPass"    ,adVarChar , adParamInput , 50 , UserPass )
		.Parameters.Append .CreateParameter( "@NewUserPass" ,adVarChar , adParamInput , 50 , NewUserPass )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub

Sub sendSmsEmail( UserId , user_email )

	Dim strFile     : strFile     = server.mapPath(FRONT_ROOT_DIR & "_skin/mail/" ) & "/mail_cPwd.html"
	Dim strTitle    : strTitle    = SEND_MAIL_NAME & " " & UserId & "님 비밀번호가 변경되었습니다."
	Dim strContants : strContants = "안전한 비밀번호 관리를 위해서 비밀번호는 반드시 정기적으로 변경을 해주시고<br>절대 다른 사람이 비밀번호를 알 수 없도록 철저히 관리해주세요.<br><br>비밀번호를 다시 변경하고자 하시면 로그인 후 비밀번호 변경 메뉴를 통하여<br>변경해주시기 바랍니다.<br><br>감사합니다."

	Dim mfrom		: mfrom		= SITE_NAME & " " & SEND_MAIL_MAIL
	Dim mto			: mto		= user_email
	Dim mtitle		: mtitle	= strTitle
	Dim mcontents	: mcontents	= ReadFile(strFile)
	mcontents	= replace(mcontents, "#USERID#", UserId )
	mcontents	= replace(mcontents, "#NOWDATE#", Now() )
	mcontents	= replace(mcontents, "#CONTANTS#", strContants )
	mcontents	= replace(mcontents, "#BOTTOM_INFO#", SEND_MAIL_BOTTOM_INFO )
	mcontents	= replace(mcontents, "#BOTTOM_COPY#", SEND_MAIL_BOTTOM_COPY )
	Dim mailMessage : mailMessage = MailSend(mtitle, mcontents, mto, mfrom, "" )

	if LEN(mailMessage) > 0 then
		response.write "메일 전송 에러 : " & mailMessage
		response.end
	end if
End Sub	
%>
<!DOCTYPE html> 
<html>
<head>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
	<title>한국외식음료협회</title>
</head>
<body>
	<script language=javascript>
		if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
		location.href = "../mypage/changePwd.asp"
	</script>
</body>
</html>