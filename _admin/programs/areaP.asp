<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/common.asp" -->
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

Dim Code        : Code        = IIF( UPLOAD__FORM("Code")="","000",UPLOAD__FORM("Code") )
Dim CodeIdx     : CodeIdx     = IIF( UPLOAD__FORM("CodeIdx")="",0,UPLOAD__FORM("CodeIdx") ) ' 프로그램Idx
Dim AddrIdx     : AddrIdx     = IIF( UPLOAD__FORM("AddrIdx")="",0,UPLOAD__FORM("AddrIdx") ) ' 지역분류Idx
Dim IntranetPwd : IntranetPwd = TagEncode( LCase( UPLOAD__FORM("IntranetPwd") ) ) ' 인트라넷 비밀번호

Dim oldFileName : oldFileName = UPLOAD__FORM("oldFileName")
Dim DellFileFg  : DellFileFg  = UPLOAD__FORM("DellFileFg")
Dim FileName    : FileName    = IIF(UPLOAD__FORM("FileName")="","",UPLOAD__FORM("FileName"))

Dim pageNo      : pageNo = UPLOAD__FORM("pageNo")

Dim sIndate     : sIndate   = UPLOAD__FORM("sIndate")
Dim sOutdate    : sOutdate  = UPLOAD__FORM("sOutdate")
Dim sName       : sName     = UPLOAD__FORM("sName")
Dim sAddr       : sAddr     = UPLOAD__FORM("sAddr")
Dim sCode       : sCode     = UPLOAD__FORM("sCode")
Dim sTel        : sTel      = UPLOAD__FORM("sTel")
Dim sPcode      : sPcode    = UPLOAD__FORM("sPcode")
Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode="     & sPcode &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sName="      & sName &_
		"&sAddr="      & sAddr &_
		"&sCode="      & sCode &_
		"&sTel="       & sTel


Call Expires()
Call dbopen()
	If actType = "INSERT" Then 

		Call Check()

		If CHECK_CNT > 0 Then 
			alertMsg = "중복된 검정장 코드 입니다. 정보를 확인해 주세요." & CHECK_CNT
		Else

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

			alertMsg = "입력되었습니다." & CHECK_CNT
		End If
		

	ElseIf actType = "UPDATE" Then 
		Call Check()

		If CHECK_CNT > 0 Then 
			alertMsg = "중복된 검정장 코드 입니다. 정보를 확인해 주세요." & CHECK_CNT
		Else
		
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
		End If

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
	"DECLARE @Name VARCHAR(200),@Addr VARCHAR(200),@Tel VARCHAR(200),@Info VARCHAR(MAX),@WebAddr VARCHAR(8000),@Map VARCHAR(200),@Code VARCHAR(10),@CodeIdx INT,@AddrIdx INT,@IntranetPwd VARCHAR(50);" &_
	"SET @Name        = ?; " &_
	"SET @Addr        = ?; " &_
	"SET @Tel         = ?; " &_
	"SET @Info        = ?; " &_
	"SET @WebAddr     = ?; " &_
	"SET @Map         = ?; " &_
	"SET @Code        = ?; " &_
	"SET @CodeIdx     = ?; " &_
	"SET @AddrIdx     = ?; " &_
	"SET @IntranetPwd = ?; " &_

	"INSERT INTO [dbo].[SP_PROGRAM_AREA]( " &_
	"	 [Name] " &_
	"	,[Addr] " &_
	"	,[Tel] " &_
	"	,[Info] " &_
	"	,[WebAddr] " &_
	"	,[Map] " &_
	"	,[Dellfg] " &_
	"	,[Code] " &_
	"	,[Indate] " &_
	"	,[CodeIdx] " &_
	"	,[AddrIdx] " &_
	"	,[IntranetPwd] " &_
	")VALUES( " &_
	"	 @Name " &_
	"	,@Addr " &_
	"	,@Tel " &_
	"	,@Info " &_
	"	,@WebAddr " &_
	"	,@Map " &_
	"	,0 " &_
	"	,@Code " &_
	"	,getDate() " &_
	"	,@CodeIdx " &_
	"	,@AddrIdx " &_
	"	,@IntranetPwd " &_
	") "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Name"        ,adVarChar     , adParamInput, 200        , Name )
		.Parameters.Append .CreateParameter( "@Addr"        ,adVarChar     , adParamInput, 200        , Addr )
		.Parameters.Append .CreateParameter( "@Tel"         ,adVarChar     , adParamInput, 200        , Tel )
		.Parameters.Append .CreateParameter( "@Info"        ,adLongVarChar , adParamInput, 2147483647 , Info )
		.Parameters.Append .CreateParameter( "@WebAddr"     ,adVarChar     , adParamInput, 8000       , WebAddr )
		.Parameters.Append .CreateParameter( "@Map"         ,adVarChar     , adParamInput, 200        , FileName )
		.Parameters.Append .CreateParameter( "@Code"        ,adVarChar     , adParamInput, 10         , Code )
		.Parameters.Append .CreateParameter( "@CodeIdx"     ,adInteger     , adParamInput, 0          , CodeIdx )
		.Parameters.Append .CreateParameter( "@AddrIdx"     ,adInteger     , adParamInput, 0          , AddrIdx )
		.Parameters.Append .CreateParameter( "@IntranetPwd" ,adVarChar     , adParamInput, 50         , IntranetPwd )
		.Execute
	End with
	call cmdclose()
End Sub
'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "set nocount on;" &_
	"DECLARE @Name VARCHAR(200),@Addr VARCHAR(200),@Tel VARCHAR(200),@Info VARCHAR(MAX),@WebAddr VARCHAR(8000),@Map VARCHAR(200),@Code VARCHAR(10),@CodeIdx INT,@AddrIdx INT, @IntranetPwd VARCHAR(50), @Idx INT;" &_
	"SET @Name        = ?; " &_
	"SET @Addr        = ?; " &_
	"SET @Tel         = ?; " &_
	"SET @Info        = ?; " &_
	"SET @WebAddr     = ?; " &_
	"SET @Map         = ?; " &_
	"SET @Code        = ?; " &_
	"SET @CodeIdx     = ?; " &_
	"SET @AddrIdx     = ?; " &_
	"SET @IntranetPwd = ?; " &_
	"SET @Idx         = ?; " &_

	"UPDATE [dbo].[SP_PROGRAM_AREA] SET " &_
	"	 [Name]       = @Name " &_
	"	,[Addr]       = @Addr " &_
	"	,[Tel]        = @Tel " &_
	"	,[Info]       = @Info " &_
	"	,[WebAddr]    = @WebAddr " &_
	"	,[Map]        = @Map " &_
	"	,[Code]       = @Code " &_
	"	,[CodeIdx]    = @CodeIdx " &_
	"	,[AddrIdx]    = @AddrIdx " &_
	"	,[IntranetPwd] = @IntranetPwd " &_
	"WHERE [Idx] = @Idx "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Name"        ,adVarChar     , adParamInput, 200        , Name )
		.Parameters.Append .CreateParameter( "@Addr"        ,adVarChar     , adParamInput, 200        , Addr )
		.Parameters.Append .CreateParameter( "@Tel"         ,adVarChar     , adParamInput, 200        , Tel )
		.Parameters.Append .CreateParameter( "@Info"        ,adLongVarChar , adParamInput, 2147483647 , Info )
		.Parameters.Append .CreateParameter( "@WebAddr"     ,adVarChar     , adParamInput, 8000       , WebAddr )
		.Parameters.Append .CreateParameter( "@Map"         ,adVarChar     , adParamInput, 200        , FileName )
		.Parameters.Append .CreateParameter( "@Code"        ,adVarChar     , adParamInput, 10         , Code )
		.Parameters.Append .CreateParameter( "@CodeIdx"     ,adInteger     , adParamInput, 0          , CodeIdx )
		.Parameters.Append .CreateParameter( "@AddrIdx"     ,adInteger     , adParamInput, 0          , AddrIdx )
		.Parameters.Append .CreateParameter( "@IntranetPwd" ,adVarChar     , adParamInput, 50         , IntranetPwd )
		.Parameters.Append .CreateParameter( "@Idx"         ,adInteger     , adParamInput, 0          , Idx )
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
	"WHERE [Idx] in( SELECT T_INT FROM @T ) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adVarChar , adParamInput, 8000 , Idx )
		.Execute
	End with
	call cmdclose()
End Sub

'체크
Sub Check()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	If actType = "INSERT" Then 
	SQL = "set nocount on;" &_
	"DECLARE @CNT INT,@Code VARCHAR(10),@Idx INT; " &_
	"SET @Code = ?; " &_
	"SET @Idx  = ?; " &_

	"SET @CNT  = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM_AREA] WHERE [Code] = @Code AND [Dellfg] = 0 ); " &_
	"SELECT @CNT AS [CNT] " 
	Else
	SQL = "set nocount on;" &_
	"DECLARE @CNT INT,@Code VARCHAR(10),@Idx INT; " &_
	"SET @Code = ?; " &_
	"SET @Idx  = ?; " &_

	"SET @CNT  = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM_AREA] WHERE [Code] = @Code AND [Idx] != @Idx AND [Dellfg] = 0 ); " &_
	"SELECT @CNT AS [CNT] " 
	End If

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Code" ,adVarChar , adParamInput, 10 , Code )
		.Parameters.Append .CreateParameter( "@Idx"  ,adInteger , adParamInput, 0  , Idx  )	
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "CHECK")
	objRs.close	: Set objRs = Nothing
End Sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN">
<html>
<head>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
</head>
<script language=javascript>
	if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
	top.location.href = "areaL.asp?<%=PageParams%>";
</script>
</html>