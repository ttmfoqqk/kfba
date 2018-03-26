<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim alertMsg : alertMsg = ""
Dim Agree1   : Agree1   = Trim( Request.Form("Agree1") )
Dim Agree2   : Agree2   = Trim( Request.Form("Agree2") )

Call Expires()
Call dbopen()
	Call AgreeProc()
	alertMsg = "정상처리 되었습니다."
Call dbclose()

Sub AgreeProc()
	SET objCmd = Server.CreateObject("adodb.command")
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @CNT INT " &_
	"SET @CNT = (SELECT COUNT(*) FROM [dbo].[SP_COMM_AGREE]) " &_
	"IF @CNT > 0  " &_
	"	BEGIN " &_
	"	UPDATE [dbo].[SP_COMM_AGREE] SET " &_
	"		 [Agree1] = ? " &_
	"		,[Agree2] = ? " &_
	"END " &_
	"ELSE " &_
	"	BEGIN " &_
	"	INSERT INTO [dbo].[SP_COMM_AGREE]( " &_
	"		 [Agree1] " &_
	"		,[Agree2] " &_
	"	)VALUES( " &_
	"		 ? " &_
	"		,? " &_
	"	) " &_
	"End  "

	call cmdopen()
	with objCmd
		.CommandText       = SQL
		.Parameters.Append .CreateParameter( "@Agree1" ,adLongVarChar  , adParamInput, 2147483647, Agree1 )
		.Parameters.Append .CreateParameter( "@Agree2" ,adLongVarChar  , adParamInput, 2147483647, Agree2 )
		.Parameters.Append .CreateParameter( "@Agree1" ,adLongVarChar  , adParamInput, 2147483647, Agree1 )
		.Parameters.Append .CreateParameter( "@Agree2" ,adLongVarChar  , adParamInput, 2147483647, Agree2 )
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
	top.location.href = "companyAgreeV.asp";
</script>
</html>