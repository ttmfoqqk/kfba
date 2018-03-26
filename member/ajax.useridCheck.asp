<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim UserId : UserId = RequestSet("UserId" , "POST" ,"")

If UserId <> "" And Len(UserId) > 4 Then
	Call Expires()
	call dbopen()
		Call getView()
	Call dbclose()
End If

Response.write FI_mCount

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT COUNT(*) AS [mCount] FROM [dbo].[SP_USER_MEMBER] WHERE [UserId] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@ID"  ,advarchar , adParamInput, 50, UserId  )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
%>