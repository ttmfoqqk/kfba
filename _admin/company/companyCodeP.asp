<!-- #include file = "../../_lib/header_utf8.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim DataMsg : DataMsg = "login"
Dim actType : actType = Trim( Request.Form("actType") )
Dim mode    : mode    = Request.Form("mode")
Dim Name    : Name    = Trim( TagEncode(Request.Form("Name")) )
Dim Ord     : Ord     = IIF( Request.Form("Ord")="",0,Request.Form("Ord") )
Dim Idx     : Idx     = Trim( Request.Form("Idx") )
Dim UsFg    : UsFg    = IIF( Request.Form("UsFg")="",0,Request.Form("UsFg") )
Dim Bigo    : Bigo    = Trim( TagEncode(Request.Form("Bigo")) )

Bigo = Replace(Bigo,vbLf,"<br>")

Call Expires()
Call dbopen()
	If mode = "1" Then 
		Call execute( "Code_p1_" & actType & "()" )
	Else
		Call execute( "Code_p2_" & actType & "()" )
	End If
	
	If err.number <> 0 Then 
		DataMsg = "error"
	Else
		DataMsg = "success"
	End If
Call dbclose()

Response.write DataMsg

Sub Code_p1_INSERT()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "" &_
	"INSERT INTO [dbo].[SP_COMM_CODE1]( " &_
	"	 [Name] " &_
	"	,[Order] " &_
	"	,[Bigo] " &_
	"	,[UsFg] " &_
	")VALUES( " &_
	"	 ? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	") "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Name" ,adVarChar , adParamInput, 50, Name )
		.Parameters.Append .CreateParameter( "@Ord"  ,adInteger , adParamInput,  0, Ord )
		.Parameters.Append .CreateParameter( "@Bigo" ,adLongVarChar , adParamInput, 2147483647, Bigo )
		.Parameters.Append .CreateParameter( "@UsFg" ,adInteger , adParamInput,  0, UsFg )
		.Execute
	End with
	call cmdclose()
End Sub

Sub Code_p1_UPDATE()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "" &_
	"UPDATE [dbo].[SP_COMM_CODE1] SET " &_
	"	 [Name]  = ? " &_
	"	,[Order] = ? " &_
	"	,[Bigo]  = ? " &_
	"	,[UsFg]  = ? " &_
	"WHERE [Idx] = ?"

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Name" ,adVarChar , adParamInput, 50, Name )
		.Parameters.Append .CreateParameter( "@Ord"  ,adInteger , adParamInput,  0, Ord )
		.Parameters.Append .CreateParameter( "@Bigo" ,adLongVarChar , adParamInput, 2147483647, Bigo )
		.Parameters.Append .CreateParameter( "@UsFg" ,adInteger , adParamInput,  0, UsFg )
		.Parameters.Append .CreateParameter( "@Idx"  ,adInteger , adParamInput,  0, Idx )
		.Execute
	End with
	call cmdclose()
End Sub

Sub Code_p1_DELETE()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "" &_
	"DECLARE @S VARCHAR (max) " &_
	"DECLARE @T TABLE(T_INT INT) " &_
	"SET @S=? " &_
	"WHILE CHARINDEX(',',@S)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(T_INT) VALUES( SUBSTRING(@S,1,CHARINDEX(',',@S)-1) ) " &_
	"	SET @S=SUBSTRING(@S,CHARINDEX(',',@S)+1,LEN(@S))  " &_
	"END " &_
	"IF CHARINDEX(',',@S)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(T_INT) VALUES( SUBSTRING(@S,1,LEN(@S)) ) " &_
	"END " &_
	"DELETE [dbo].[SP_COMM_CODE1] WHERE [Idx] in(SELECT T_INT FROM @T) " &_
	"DELETE [dbo].[SP_COMM_CODE2] WHERE [PIdx] in(SELECT T_INT FROM @T) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"  ,adVarChar , adParamInput,  8000, Idx )
		.Execute
	End with
	call cmdclose()
End Sub

Sub Code_p2_INSERT()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "" &_
	"INSERT INTO [dbo].[SP_COMM_CODE2]( " &_
	"	 [PIdx] " &_
	"	,[Name] " &_
	"	,[Order] " &_
	"	,[Bigo] " &_
	"	,[UsFg] " &_
	")VALUES( " &_
	"	 ? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	") "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"  ,adInteger , adParamInput,  0, Idx )
		.Parameters.Append .CreateParameter( "@Name" ,adVarChar , adParamInput, 50, Name )		
		.Parameters.Append .CreateParameter( "@Ord"  ,adInteger , adParamInput,  0, Ord )
		.Parameters.Append .CreateParameter( "@Bigo" ,adLongVarChar , adParamInput, 2147483647, Bigo )
		.Parameters.Append .CreateParameter( "@UsFg" ,adInteger , adParamInput,  0, UsFg )
		.Execute
	End with
	call cmdclose()
End Sub

Sub Code_p2_UPDATE()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "" &_
	"UPDATE [dbo].[SP_COMM_CODE2] SET " &_
	"	 [Name]  = ? " &_
	"	,[Order] = ? " &_
	"	,[Bigo]  = ? " &_
	"	,[UsFg]  = ? " &_
	"WHERE [Idx] = ?"

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Name" ,adVarChar , adParamInput, 50, Name )
		.Parameters.Append .CreateParameter( "@Ord"  ,adInteger , adParamInput,  0, Ord )
		.Parameters.Append .CreateParameter( "@Bigo" ,adLongVarChar , adParamInput, 2147483647, Bigo )
		.Parameters.Append .CreateParameter( "@UsFg" ,adInteger , adParamInput,  0, UsFg )
		.Parameters.Append .CreateParameter( "@Idx"  ,adInteger , adParamInput,  0, Idx )
		.Execute
	End with
	call cmdclose()
End Sub

Sub Code_p2_DELETE()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "" &_
	"DECLARE @S VARCHAR (max) " &_
	"DECLARE @T TABLE(T_INT INT) " &_
	"SET @S=? " &_
	"WHILE CHARINDEX(',',@S)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(T_INT) VALUES( SUBSTRING(@S,1,CHARINDEX(',',@S)-1) ) " &_
	"	SET @S=SUBSTRING(@S,CHARINDEX(',',@S)+1,LEN(@S))  " &_
	"END " &_
	"IF CHARINDEX(',',@S)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(T_INT) VALUES( SUBSTRING(@S,1,LEN(@S)) ) " &_
	"END " &_
	"DELETE [dbo].[SP_COMM_CODE2] WHERE [Idx] in(SELECT T_INT FROM @T) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"  ,adVarChar , adParamInput,  8000, Idx )
		.Execute
	End with
	call cmdclose()
End Sub
%>