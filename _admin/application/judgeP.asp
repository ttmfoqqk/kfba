<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim alertMsg   : alertMsg   = ""

Dim actType    : actType    = RequestSet("actType" ,"POST","")
Dim Idx        : Idx        = RequestSet("Idx" ,"POST","")
Dim State      : State      = RequestSet("State" ,"POST",1)

Dim pageNo     : pageNo    = RequestSet("pageNo" , "POST" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate" , "POST" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate", "POST" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "POST" , "")
Dim sState     : sState     = RequestSet("sState"    , "POST" , "")
Dim sId        : sId        = RequestSet("sId"    , "POST" , "")
Dim sName      : sName      = RequestSet("sName"    , "POST" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"    , "POST" , "")
Dim sBirth     : sBirth     = RequestSet("sBirth"    , "POST" , "")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sPcode="     & sPcode &_
		"&sState="     & sState &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sBirth="     & sBirth


Call Expires()
Call dbopen()
	If actType = "UPDATE" Then 
		Call Update()
		alertMsg = "수정되었습니다."
	ElseIf actType = "DELETE" Then 
		Call Delete()
		alertMsg = "삭제되었습니다."
	Else
		alertMsg = "[actType] 이 없습니다."
	End If
	
Call dbclose()


'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "UPDATE [dbo].[SP_PROGRAM_JUDGE_APP] SET " &_
	"	 [State]    = ? " &_
	"WHERE [Idx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@State" ,adInteger , adParamInput, 0 , State )
		.Parameters.Append .CreateParameter( "@Idx"   ,adInteger , adParamInput, 0 , Idx )
		.Execute
	End with
	call cmdclose()
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
	
	
	"UPDATE [dbo].[SP_PROGRAM_JUDGE_APP] SET " &_
	"	[Dellfg] = 1 " &_
	"WHERE [Idx] in( SELECT T_INT FROM @T ) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adVarChar , adParamInput, 8000 , Idx )
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
	top.location.href = "judgeL.asp?<%=PageParams%>";
</script>
</html>