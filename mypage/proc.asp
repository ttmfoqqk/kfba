<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim savePath : savePath = "\board/" '첨부 저장경로.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath
UPLOAD__FORM.MaxFileLen		= 10 * 1024 * 1024 '10메가

Dim alertMsg : alertMsg = ""
Dim Idx      : Idx            = IIF( UPLOAD__FORM("Idx")="","",UPLOAD__FORM("Idx") )
Dim pageNo   : pageNo         = IIF( UPLOAD__FORM("pageNo")="","1",UPLOAD__FORM("pageNo") )
Dim sName    : sName          = UPLOAD__FORM("sName")
Dim sId      : sId            = UPLOAD__FORM("sId")
Dim sTitle   : sTitle         = UPLOAD__FORM("sTitle")
Dim sContant : sContant       = UPLOAD__FORM("sContant")
Dim BoardKey : BoardKey       = IIF( UPLOAD__FORM("BoardKey")="","",UPLOAD__FORM("BoardKey") )
Dim actType  : actType        = UPLOAD__FORM("actType")

Dim oldFileName : oldFileName = UPLOAD__FORM("oldFileName")
Dim Title       : Title       = IIF(UPLOAD__FORM("Title")="","",TagEncode( UPLOAD__FORM("Title") ))
Dim Contants    : Contants    = IIF(UPLOAD__FORM("Contants")="","",TagEncode(UPLOAD__FORM("Contants")) )
Dim FileName    : FileName    = IIF(UPLOAD__FORM("FileName")="","",UPLOAD__FORM("FileName"))
Dim DellFileFg  : DellFileFg  = UPLOAD__FORM("DellFileFg")
Dim Notice      : Notice      = IIF(UPLOAD__FORM("Notice")="",0,UPLOAD__FORM("Notice"))

Dim Secret      : Secret      = IIF(UPLOAD__FORM("Secret")="",0,UPLOAD__FORM("Secret"))
Dim Pwd         : Pwd         = IIF(UPLOAD__FORM("Pwd")="","",UPLOAD__FORM("Pwd"))

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&BoardKey=" & BoardKey &_
		"&sName="    & sName &_
		"&sId="      & sId &_
		"&sTitle="   & sTitle &_
		"&sContant=" & sContant &_
		"&sWord="    & sWord

Call Expires()
Call dbopen()
	Call BoardCodeView()
	If BDV_Idx = "" Or BDV_State = "1" Then
		
		Call dbclose()
		With Response
		 .Write "<script language='javascript' type='text/javascript'>"
		 .Write "alert('잘못된 게시판 코드 입니다.');"
		 .Write "history.back(-1);"
		 .Write "</script>"
		 .End
		End With
	End If

	if (alertMsg <> "")	then
		actType	= ""
	Elseif (actType = "INSERT") Then	'글작성
		
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
		alertMsg = "입력 되었습니다."
	Elseif (actType = "ANSWERE") Then	'글작성
		
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
		
		Call Answere()
		alertMsg = "입력 되었습니다."

	ElseIf (actType = "MODIFY") Then	'글수정
		
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
		alertMsg = "수정 되었습니다."
	ElseIf (actType = "DELETE") Then	'글삭제
		
		'글 삭제시 파일 삭제
		'If FI_FileName <> "" Then
		'	Set FSO = CreateObject("Scripting.FileSystemObject")
		'		If (FSO.FileExists(ETING_UPLOAD_BASE_PATH & savePath & FI_File_name)) Then	' 파일삭제
		'			fso.deletefile(ETING_UPLOAD_BASE_PATH & savePath & FI_File_name)
		'		End If
		'	set FSO = Nothing
		'End If

		Call Delete()
		alertMsg = "삭제 되었습니다."
	else
		alertMsg = "actType[" & actType & "]이 정의되지 않았습니다."
	end If
	
Call dbclose()

'입력
Sub Insert()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_

	"INSERT INTO [dbo].[SP_BOARD]" &_
	"( [BoardKey],[Title],[Contants],[File],[Secret],[Pwd],[Notice],[Order],[Depth],[Parent],[UserIdx],[RCnt],[CmCnt],[Dellfg],[Indate],[Ip] )" &_
	"VALUES" &_
	"( ?         ,?      ,?         ,?     ,?       ,?    ,?       ,0      ,0      ,0       ,?         ,0     ,0      ,0       ,getDate(),?   );" &_
	
	"UPDATE [dbo].[SP_BOARD] SET [Parent] = [Idx] WHERE [Idx] = SCOPE_IDENTITY(); "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@BoardKey" ,adInteger     , adParamInput, 50         , BoardKey )
		.Parameters.Append .CreateParameter( "@Title"    ,adVarChar     , adParamInput, 50         , Title )
		.Parameters.Append .CreateParameter( "@Contants" ,adLongVarChar , adParamInput, 2147483647 , Contants )
		.Parameters.Append .CreateParameter( "@File"     ,adVarChar     , adParamInput, 200        , FileName )
		.Parameters.Append .CreateParameter( "@Secret"   ,adInteger     , adParamInput, 0          , Secret )
		.Parameters.Append .CreateParameter( "@Pwd"      ,adVarChar     , adParamInput, 50         , Pwd )
		.Parameters.Append .CreateParameter( "@Notice"   ,adInteger     , adParamInput, 0          , Notice )
		.Parameters.Append .CreateParameter( "@UserIdx"  ,adInteger     , adParamInput, 0          , Session("UserIdx") )
		.Parameters.Append .CreateParameter( "@Ip"       ,adVarChar     , adParamInput, 20         , g_uip )
		.Execute
	End with
	call cmdclose()
End Sub
'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "UPDATE [dbo].[SP_BOARD] SET " &_
	"	 [Title]    = ? " &_
	"	,[Contants] = ? " &_
	"	,[File]     = ? " &_
	"	,[Secret]   = ? " &_
	"WHERE [Idx] = ? AND [UserIdx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Title"    ,adVarChar     , adParamInput, 50         , Title )
		.Parameters.Append .CreateParameter( "@Contants" ,adLongVarChar , adParamInput, 2147483647 , Contants )
		.Parameters.Append .CreateParameter( "@File"     ,adVarChar     , adParamInput, 200        , FileName )
		.Parameters.Append .CreateParameter( "@Secret"   ,adInteger     , adParamInput, 0          , Secret )
		.Parameters.Append .CreateParameter( "@Idx"      ,adInteger     , adParamInput, 0          , Idx )
		.Parameters.Append .CreateParameter( "@UserIdx"  ,adInteger     , adParamInput, 0          , Session("UserIdx") )
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
	
	
	"UPDATE [dbo].[SP_BOARD] SET " &_
	"	[Dellfg] = 1 " &_
	"WHERE ( [Idx] in( SELECT T_INT FROM @T ) OR [Parent] in( SELECT T_INT FROM @T ) ) "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "Idx" ,adVarChar , adParamInput, 8000 , Idx )
		.Execute
	End with
	call cmdclose()
End Sub
%>
<!DOCTYPE html>
<HTML>
<HEAD>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
</head>
<script language=javascript>
	if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
	top.location.href = "list.asp?<%=PageParams%>";
</script>
</html>