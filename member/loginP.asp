<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
If Session("UserIdx") <> "" Then 
	Response.redirect "../mypage/"
End If

Dim GoUrl   : GoUrl   = RequestSet("GoUrl"   ,"POST" ,FRONT_ROOT_DIR)
Dim UserId  : UserId  = RequestSet("UserId"  ,"POST" ,"")
Dim UserPwd : UserPwd = RequestSet("UserPwd" ,"POST" ,"")
Dim SaveLog : SaveLog = RequestSet("SaveLog" ,"POST" ,"")

if UserId="" Or  UserPwd = "" Then
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('입력하신 아이디 혹은 비밀번호가 일치하지 않습니다.\n\n대소문자 확인 후 입력해주세요!');"
	 .Write "history.go(-1);"
	 .Write "</script>"
	 .End
	End With
End If


Call Expires()
Call dbopen()
	Call Check()

	If FV_UserPass = "1" Then

		' 쿠키 생성
		If SaveLog = "Y" Then 
			response.cookies("UserIdSave")("id")    = UserId
			response.cookies("UserIdSave")("pwd")   = UserPwd
			response.cookies("UserIdSave")("check") = "Y"
			Response.Cookies("UserIdSave").domain   = Request.ServerVariables("SERVER_NAME")
			response.cookies("UserIdSave").expires  = Now() + 365
		Else
			response.cookies("UserIdSave").expires  = Now() - 1
		End If

		Session("UserIdx")  = FV_UserIdx
		Session("UserId")   = FV_UserId
		Session("UserName") = FV_UserName

		response.redirect GoUrl
	Else
		With Response
		 .Write "<script language='javascript' type='text/javascript'>"
		 .Write "alert('입력하신 아이디 혹은 비밀번호가 일치하지 않습니다.\n\n대소문자 확인 후 입력해주세요!');"
		 .Write "history.go(-1);"
		 .Write "</script>"
		 .End
		End With
	End If

Call dbclose()

'로그인 조회
Sub Check()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT [UserIdx] , [UserId] , [UserName] , pwdcompare( ? ,[UserPass] ) as [UserPass] "  &_
	" FROM [dbo].[SP_USER_MEMBER] WHERE [UserId] = ? AND [UserDelFg] = 0 "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@PWD" ,advarchar , adParamInput, 50, UserPwd )
		.Parameters.Append .CreateParameter( "@ID"  ,advarchar , adParamInput, 50, UserId  )		
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")
	Set objRs = Nothing
End Sub
%>