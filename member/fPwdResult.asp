<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
If Session("UserIdx") <> "" Then 
	Response.redirect "../mypage/"
End If
Dim NEW_PASSWORD : NEW_PASSWORD = RandomNumber(10,"")

Dim ResultMsg  : ResultMsg  = "�Է��Ͻ� ������ ��ġ�ϴ� ���̵� �����ϴ�.<br>��Ȯ�� ������ Ȯ�� �� �ٽ� �Է� ��Ź �帳�ϴ�."

If session("sName") <> "" and session("sDupInfo") <> "" Then 
	Call Expires()
	Call dbopen()
		Call Check()
		If FI_UserIdx <> "" Then 
			'ResultMsg="������ �ӽ� ��й�ȣ��  <font color='ff7ebb'>" & FI_UserEmail & "</font> ���� <br>�߼��߽��ϴ�.<br>�ӽ� ��й�ȣ : [ <span style='color:#ff469d;cursor:pointer' onclick=""TextClipBoard('" & Trim(NEW_PASSWORD) & "')"">" & Trim(NEW_PASSWORD) & "</span> ]"
			ResultMsg="������ �ӽ� ��й�ȣ : [ <span style='cursor:pointer;font-weight:bold;' onclick=""TextClipBoard('" & Trim(NEW_PASSWORD) & "')"">" & Trim(NEW_PASSWORD) & "</span> ]"

			Call RandomPwUpdate( FI_UserIdx ) ' �ӽú�й�ȣ�� ��ü
			'Call sendSmsEmail( FI_UserId , FI_UserEmail ) ' ���� �߼�

			session("sVNumber")      = ""
			session("sName")         = ""
			session("sBirthDate")    = ""
			session("sGender")       = ""
			session("sNationalInfo") = ""
			session("sDupInfo")      = ""
			session("sConnInfo")     = ""

		End If
	Call dbclose() 
End If

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "member/fPwdResult.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("ResultMsg", ResultMsg ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = Nothing

'-----------------------------------------------
' ��й�ȣ��ȸ
'-----------------------------------------------
Sub Check()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SELECT top 1 UserEmail , UserId ,UserIdx "  &_
	" FROM [dbo].[SP_USER_MEMBER] "  &_
	" WHERE [UserName] = ? "  &_
	" AND [UserDIKEY] = ? "  &_
	" AND UserDelfg = 0 ORDER BY [UserIdx] DESC "

   
	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( ,advarchar , adParamInput,   50, session("sName") )
		.Parameters.Append .CreateParameter( ,advarchar , adParamInput,   64, session("sDupInfo")  )
		set objRs = .Execute
	End with
	call cmdclose()

	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
	
End Sub

'-----------------------------------------------
' ��й�ȣ ������ ��ü
'-----------------------------------------------
Sub RandomPwUpdate(UserIdx)
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "UPDATE [dbo].[SP_USER_MEMBER] SET"  &_
	" UserPass = pwdencrypt(?) "  &_
	" WHERE [UserIdx] = ? "
	
	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( ,advarchar , adParamInput,   128, NEW_PASSWORD )
		.Parameters.Append .CreateParameter( ,adInteger , adParamInput,   0, UserIdx )
		.Execute
	End with
	call cmdclose()
End Sub

'-----------------------------------------------
' Email ����
'-----------------------------------------------
Sub sendSmsEmail(UserId, user_email)
	Dim strFile     : strFile     = server.mapPath(FRONT_ROOT_DIR & "_skin/mail/" ) & "/mail_fPwd.html"
	Dim strTitle    : strTitle    = SEND_MAIL_NAME & " " & UserId & "��û�Ͻ� ��й�ȣ �Դϴ�."
	Dim strContants : strContants = "�α��� �� ���ο� ��й�ȣ�� �����Ͽ� �̿��Ͻñ� �ٶ��ϴ�."

	Dim mfrom		: mfrom		= SITE_NAME & " " & SEND_MAIL_MAIL
	Dim mto			: mto		= user_email
	Dim mtitle		: mtitle	= strTitle
	Dim mcontents	: mcontents	= ReadFile( strFile )
		mcontents	= replace(mcontents, "#USERID#", StrLenBlind(UserId,2) )
		mcontents	= replace(mcontents, "#PASSWORD#", NEW_PASSWORD )
		mcontents	= replace(mcontents, "#NOWDATE#", Now() )
		mcontents	= replace(mcontents, "#CONTANTS#", strContants )
		mcontents	= replace(mcontents, "#BOTTOM_INFO#", SEND_MAIL_BOTTOM_INFO )
		mcontents	= replace(mcontents, "#BOTTOM_COPY#", SEND_MAIL_BOTTOM_COPY )
	Dim mailMessage : mailMessage = MailSend(mtitle, mcontents, mto, mfrom, attachPath)
	if LEN(mailMessage) > 0 then
		'response.write "���� ���� ���� : " & mailMessage
		'response.end
	end if
	
End Sub	
%>