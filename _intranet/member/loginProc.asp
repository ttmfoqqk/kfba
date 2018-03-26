<!-- #include file = "../../_lib/header.asp" -->
<%
Dim AdminId  : AdminId  = Trim( request("AdminId") )
Dim pass     : pass     = Trim( request("pass") )
Dim GoUrl    : GoUrl    = request("GoUrl")
Dim firstURL : firstURL = IIF( GoUrl="" , "../application/applicationL.asp" , GoUrl )

If AdminId = "" Or pass = "" Then
	Call msgbox("입력하신 아이디 혹은 비밀번호가 일치하지 않습니다."&vbCrLf & vbCrLf&"대소문자 확인 후 입력해주세요!1", true)
End If

Call Expires()
Call dbopen()
	Call Check()
Call dbclose()

If FV_Idx <> "" AND FV_Code <> "" Then
	Session("Judge_Idx")  = FV_Idx
	Session("Judge_Id")   = FV_Code
	Session("Judge_Name") = FV_Name

	response.redirect firstURL
Else

	Call msgbox("입력하신 아이디 혹은 비밀번호가 일치하지 않습니다."&vbCrLf & vbCrLf&"대소문자 확인 후 입력해주세요!2", true)

End If

Sub Check()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @ID VARCHAR(10), @PWD VARCHAR(50);" &_
	"SET @ID  = ? " &_
	"SET @PWD = ? " &_
	
	"SELECT TOP 1 " &_
	"	 [Idx] " &_
	"	,[Name] " &_
	"	,[Code] " &_
	"FROM [dbo].[SP_PROGRAM_AREA] " &_
	"WHERE [Code] = @ID " &_
	"AND [IntranetPwd] = @PWD " &_
	"AND [Dellfg] = 0 "


	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@ID"  ,advarchar , adParamInput,   10, AdminId  )
		.Parameters.Append .CreateParameter( "@PWD" ,advarchar , adParamInput,   50, pass )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")
	Set objRs = Nothing
End Sub
%>