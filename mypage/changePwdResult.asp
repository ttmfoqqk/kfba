<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
checkLogin( g_host & g_url )
Dim UserPass      : UserPass      = RequestSet("UserPass" , "POST" , "")
Dim NewUserPass   : NewUserPass   = RequestSet("NewUserPass" , "POST" , "")
Dim UserPassCheck : UserPassCheck = RequestSet("UserPassCheck" , "POST" , "")
Dim alertMsg      : alertMsg = "�����Ǿ����ϴ�."

If UserPass="" Or NewUserPass="" Or UserPassCheck="" Then 
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('ERR : �߸��� ��� �Դϴ�.');"
	 .Write "history.go(-1);"
	 .Write "</script>"
	 .End
	End With
End If

If NewUserPass <> UserPassCheck Then 
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('ERR : �� ��й�ȣ�� �߸��ԷµǾ����ϴ�.');"
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
	'Call sendSmsEmail( FI_USERID , FI_USERMAIL ) ' ���� �߼�
Else
	alertMsg = "ERR : ��й�ȣ�� ��Ȯ���� �ʽ��ϴ�."
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
	Dim strTitle    : strTitle    = SEND_MAIL_NAME & " " & UserId & "�� ��й�ȣ�� ����Ǿ����ϴ�."
	Dim strContants : strContants = "������ ��й�ȣ ������ ���ؼ� ��й�ȣ�� �ݵ�� ���������� ������ ���ֽð�<br>���� �ٸ� ����� ��й�ȣ�� �� �� ������ ö���� �������ּ���.<br><br>��й�ȣ�� �ٽ� �����ϰ��� �Ͻø� �α��� �� ��й�ȣ ���� �޴��� ���Ͽ�<br>�������ֽñ� �ٶ��ϴ�.<br><br>�����մϴ�."

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
		response.write "���� ���� ���� : " & mailMessage
		response.end
	end if
End Sub	
%>
<!DOCTYPE html> 
<html>
<head>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
	<title>�ѱ��ܽ�������ȸ</title>
</head>
<body>
	<script language=javascript>
		if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
		location.href = "../mypage/changePwd.asp"
	</script>
</body>
</html>