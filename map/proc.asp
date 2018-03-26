<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim savePath : savePath = "\programsArea/" '첨부 저장경로.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath
UPLOAD__FORM.MaxFileLen		= 10 * 1024 * 1024 '10메가

Dim alertMsg    : alertMsg    = ""

Dim actType     : actType     = UPLOAD__FORM("actType")
Dim Idx         : Idx         = IIF(UPLOAD__FORM("Idx")="",0,UPLOAD__FORM("Idx"))
Dim Name        : Name        = TagEncode( IIF(UPLOAD__FORM("Name")="","",UPLOAD__FORM("Name")) )
Dim Addr        : Addr        = TagEncode( IIF(UPLOAD__FORM("Addr")="","",UPLOAD__FORM("Addr")) )
Dim Tel         : Tel         = TagEncode( IIF(UPLOAD__FORM("Tel")="","",UPLOAD__FORM("Tel")) )
Dim Info        : Info        = IIF(UPLOAD__FORM("Info")="","",UPLOAD__FORM("Info"))
Dim WebAddr     : WebAddr     = TagEncode( IIF(UPLOAD__FORM("WebAddr")="","",UPLOAD__FORM("WebAddr")) )

Dim CodeIdx     : CodeIdx     = IIF( UPLOAD__FORM("CodeIdx")="",0,UPLOAD__FORM("CodeIdx") ) ' 프로그램Idx
Dim AddrIdx     : AddrIdx     = IIF( UPLOAD__FORM("AddrIdx")="",0,UPLOAD__FORM("AddrIdx") ) ' 지역분류Idx

Dim oldFileName : oldFileName = UPLOAD__FORM("oldFileName")
Dim DellFileFg  : DellFileFg  = UPLOAD__FORM("DellFileFg")
Dim FileName    : FileName    = IIF(UPLOAD__FORM("FileName")="","",UPLOAD__FORM("FileName"))

Dim pageNo      : pageNo = UPLOAD__FORM("pageNo")
Dim sPcode      : sPcode = UPLOAD__FORM("sPcode")
Dim sACode      : sACode = UPLOAD__FORM("sACode")
Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode=" & sPcode &_
		"&sACode=" & sACode


Call Expires()
Call dbopen()
	If actType = "INSERT" Then 

		If FileName <>"" Then 
			If FILE_CHECK_EXT(FileName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("FileName").FileLen Then 
					FileName = DextFileUpload("FileName",UPLOAD_BASE_PATH & savePath,0)
				Else
					With Response
					 .Write "<script language='javascript' type='text/javascript'>"
					 .Write "alert('파일의 크기는 10MB 를 넘길수 없습니다.');"
					 .Write "history.go(-1);"
					 .Write "</script>"
					 .End
					End With
				End If
			Else
				With Response
				 .Write "<script language='javascript' type='text/javascript'>"
				 .Write "alert('잘못된 파일입니다. [asp,php,jsp,html,js] 파일은 업로드 할수 없습니다.');"
				 .Write "history.go(-1);"
				 .Write "</script>"
				 .End
				End With
			End If
		End If

		Call Insert()
		alertMsg = "입력되었습니다."
	ElseIf actType = "UPDATE" Then 
		
		If FileName <>"" Then 
			If FILE_CHECK_EXT(FileName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("FileName").FileLen Then 
					FileName = DextFileUpload("FileName",UPLOAD_BASE_PATH & savePath,0)
				Else
					With Response
					 .Write "<script language='javascript' type='text/javascript'>"
					 .Write "alert('파일의 크기는 10MB 를 넘길수 없습니다.');"
					 .Write "history.go(-1);"
					 .Write "</script>"
					 .End
					End With
				End If
			Else
				With Response
				 .Write "<script language='javascript' type='text/javascript'>"
				 .Write "alert('잘못된 파일입니다. [asp,php,jsp,html,js] 파일은 업로드 할수 없습니다.');"
				 .Write "history.go(-1);"
				 .Write "</script>"
				 .End
				End With
			End If

			If oldFileName <> "" Then
				Set FSO = CreateObject("Scripting.FileSystemObject")
					If (FSO.FileExists(UPLOAD_BASE_PATH & savePath & oldFileName)) Then	' 같은 이름의 파일이 있을 때 삭제
						fso.deletefile(UPLOAD_BASE_PATH & savePath & oldFileName)
					End If
				set FSO = Nothing
			End If
		Else
			FileName = oldFileName
		End If

		If DellFileFg = "1" Then 
			If oldFileName <> "" Then
				Set FSO = CreateObject("Scripting.FileSystemObject")
					If (FSO.FileExists(UPLOAD_BASE_PATH & savePath & oldFileName)) Then	' 같은 이름의 파일이 있을 때 삭제
						fso.deletefile(UPLOAD_BASE_PATH & savePath & oldFileName)
					End If
				set FSO = Nothing
			End If

			FileName = ""
		End If

		Call Update()
		alertMsg = "수정되었습니다."
	ElseIf actType = "DELETE" Then 
		Call Delete()
		alertMsg = "삭제되었습니다."
	Else
		alertMsg = "[actType] 이 없습니다."
	End If
	
Call dbclose()

'입력
Sub Insert()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "set nocount on;" &_
	"INSERT INTO [dbo].[SP_PROGRAM_AREA]( " &_
	"	 [Name] " &_
	"	,[Addr] " &_
	"	,[Tel] " &_
	"	,[Info] " &_
	"	,[WebAddr] " &_
	"	,[Map] " &_
	"	,[Dellfg] " &_
	"	,[Indate] " &_
	"	,[CodeIdx] " &_
	"	,[AddrIdx] " &_
	"	,[UserIdx] " &_
	")VALUES( " &_
	"	 ? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,0 " &_
	"	,getDate() " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	") "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Name"    ,adVarChar     , adParamInput, 200        , Name )
		.Parameters.Append .CreateParameter( "@Addr"    ,adVarChar     , adParamInput, 200        , Addr )
		.Parameters.Append .CreateParameter( "@Tel"     ,adVarChar     , adParamInput, 200        , Tel )
		.Parameters.Append .CreateParameter( "@Info"    ,adLongVarChar , adParamInput, 2147483647 , Info )
		.Parameters.Append .CreateParameter( "@WebAddr" ,adVarChar     , adParamInput, 8000       , WebAddr )
		.Parameters.Append .CreateParameter( "@Map"     ,adVarChar     , adParamInput, 200        , FileName )
		.Parameters.Append .CreateParameter( "@CodeIdx" ,adInteger     , adParamInput, 0          , CodeIdx )
		.Parameters.Append .CreateParameter( "@AddrIdx" ,adInteger     , adParamInput, 0          , AddrIdx )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger     , adParamInput, 0          , IIF(session("UserIdx")="",0,session("UserIdx")) )
		.Execute
	End with
	call cmdclose()
End Sub
'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "UPDATE [dbo].[SP_PROGRAM_AREA] SET " &_
	"	 [Name]    = ? " &_
	"	,[Addr]    = ? " &_
	"	,[Tel]     = ? " &_
	"	,[Info]    = ? " &_
	"	,[WebAddr] = ? " &_
	"	,[Map]     = ? " &_
	"	,[CodeIdx] = ? " &_
	"	,[AddrIdx] = ? " &_
	"WHERE [Idx] = ? AND [UserIdx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Name"    ,adVarChar     , adParamInput, 200        , Name )
		.Parameters.Append .CreateParameter( "@Addr"    ,adVarChar     , adParamInput, 200        , Addr )
		.Parameters.Append .CreateParameter( "@Tel"     ,adVarChar     , adParamInput, 200        , Tel )
		.Parameters.Append .CreateParameter( "@Info"    ,adLongVarChar , adParamInput, 2147483647 , Info )
		.Parameters.Append .CreateParameter( "@WebAddr" ,adVarChar     , adParamInput, 8000       , WebAddr )
		.Parameters.Append .CreateParameter( "@Map"     ,adVarChar     , adParamInput, 200        , FileName )
		.Parameters.Append .CreateParameter( "@CodeIdx" ,adInteger     , adParamInput, 0          , CodeIdx )
		.Parameters.Append .CreateParameter( "@AddrIdx" ,adInteger     , adParamInput, 0          , AddrIdx )
		.Parameters.Append .CreateParameter( "@Idx"     ,adInteger     , adParamInput, 0          , Idx )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger     , adParamInput, 0          , IIF(session("UserIdx")="",0,session("UserIdx")) )
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
	
	
	"UPDATE [dbo].[SP_PROGRAM_AREA] SET " &_
	"	[Dellfg] = 1 " &_
	"WHERE [Idx] in( SELECT T_INT FROM @T ) AND [UserIdx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adVarChar , adParamInput, 8000 , Idx )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger     , adParamInput, 0          , IIF(session("UserIdx")="",0,session("UserIdx")) )
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
	top.location.href = "write.asp";
</script>
</html>