<!-- #include file = "../../_lib/header.asp" -->
<%
Dim AdminId  : AdminId  = request("AdminId")
Dim pass     : pass     = request("pass")
Dim GoUrl    : GoUrl    = request("GoUrl")
Dim firstURL : firstURL = IIF( GoUrl="" , "../company/companyAgreeV.asp" , GoUrl )

Call Expires()
Call dbopen()
	Call Check()
Call dbclose()

If FI_Idx <> "" Then 
	Session("Admin_Idx")  = FI_Idx
	Session("Admin_Id")   = FI_Id
	Session("Admin_Name") = FI_Name
	response.redirect firstURL
Else
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('로그인실패.');"
	 .Write "history.back(-1);"
	 .Write "</script>"
	 .End
	End With
End If

Sub Check()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT [Idx] , [Id] , [Name] "  &_
	" FROM [dbo].[SP_ADMIN_MEMBER] WHERE [Id] = ? "  &_
	" AND [Pwd] = ? "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@ID" ,advarchar , adParamInput,   50, AdminId  )
		.Parameters.Append .CreateParameter( "@PWD" ,advarchar , adParamInput,   50, pass )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
%>