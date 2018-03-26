<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim savePathJob   : savePathJob = "\job/" '첨부 저장경로.
Dim savePathPhoto : savePathPhoto = "\appMember/" '첨부 저장경로.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePathJob
UPLOAD__FORM.MaxFileLen		= 50 * 1024 * 1024 '10메가

Dim alertMsg       : alertMsg = ""
Dim actType        : actType        = UPLOAD__FORM("actType")

Dim Idx            : Idx            = UPLOAD__FORM("Idx")
Dim Form           : Form           = IIF( UPLOAD__FORM("Form")="" , 0 , UPLOAD__FORM("Form") )
Dim Kind           : Kind           = IIF( UPLOAD__FORM("Kind")="" , 0 , UPLOAD__FORM("Kind") )
Dim WorkArea       : WorkArea       = UPLOAD__FORM("WorkArea")
Dim Pay            : Pay            = UPLOAD__FORM("Pay")
Dim School         : School         = UPLOAD__FORM("School")
Dim Bigo           : Bigo           = UPLOAD__FORM("Bigo")
Dim PhotoName      : PhotoName      = UPLOAD__FORM("PhotoName")
Dim FileName       : FileName       = UPLOAD__FORM("FileName")
Dim oldPhotoName   : oldPhotoName   = UPLOAD__FORM("oldPhotoName")
Dim oldFileName    : oldFileName    = UPLOAD__FORM("oldFileName")


Dim careerName     : careerName     = UPLOAD__FORM("careerName")
Dim careerMonth    : careerMonth    = UPLOAD__FORM("careerMonth")
Dim careerPosition : careerPosition = UPLOAD__FORM("careerPosition")

Dim qualifyName    : qualifyName    = UPLOAD__FORM("qualifyName")
Dim qualifyDate    : qualifyDate    = UPLOAD__FORM("qualifyDate")
Dim qualifyPublish : qualifyPublish = UPLOAD__FORM("qualifyPublish")

Dim pageNo         : pageNo         = UPLOAD__FORM("pageNo")
Dim sName          : sName          = UPLOAD__FORM("sName")
Dim sId            : sId            = UPLOAD__FORM("sId")
Dim sTitle         : sTitle         = UPLOAD__FORM("sTitle")
Dim sContant       : sContant       = UPLOAD__FORM("sContant")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sName="    & sName &_
		"&sId="      & sId &_
		"&sTitle="   & sTitle &_
		"&sContant=" & sContant &_
		"&sWord="    & sWord


Call Expires()
Call dbopen()

	if (alertMsg <> "")	then
		actType	= ""
	Elseif (actType = "INSERT") Then	'글작성
		If FileName <>"" Then 
			If FILE_CHECK_EXT(FileName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("FileName").FileLen Then 
					FileName = DextFileUpload("FileName",UPLOAD_BASE_PATH & savePathJob,0)
				Else
					With Response
					 .Write "<script language='javascript' type='text/javascript'>"
					 .Write "alert('파일의 크기는 50MB 를 넘길수 없습니다.');"
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
		'' 사진
		If PhotoName <>"" Then 
			If FILE_CHECK_EXT(PhotoName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("PhotoName").FileLen Then 
					PhotoName = DextFileUpload("PhotoName",UPLOAD_BASE_PATH & savePathPhoto,0)
				Else
					With Response
					 .Write "<script language='javascript' type='text/javascript'>"
					 .Write "alert('파일의 크기는 50MB 를 넘길수 없습니다.');"
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
	
	ElseIf (actType = "MODIFY") Then	'글수정
		If FileName <>"" Then 
			If FILE_CHECK_EXT(FileName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("FileName").FileLen Then 
					FileName = DextFileUpload("FileName",UPLOAD_BASE_PATH & savePathJob,0)
				Else
					With Response
					 .Write "<script language='javascript' type='text/javascript'>"
					 .Write "alert('파일의 크기는 50MB 를 넘길수 없습니다.');"
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
					If (FSO.FileExists(UPLOAD_BASE_PATH & savePathJob & oldFileName)) Then	' 같은 이름의 파일이 있을 때 삭제
						fso.deletefile(UPLOAD_BASE_PATH & savePathJob & oldFileName)
					End If
				set FSO = Nothing
			End If
		Else
			FileName = oldFileName
		End If

		If DellFileFg = "1" Then 
			If oldFileName <> "" Then
				Set FSO = CreateObject("Scripting.FileSystemObject")
					If (FSO.FileExists(UPLOAD_BASE_PATH & savePathJob & oldFileName)) Then	' 같은 이름의 파일이 있을 때 삭제
						fso.deletefile(UPLOAD_BASE_PATH & savePathJob & oldFileName)
					End If
				set FSO = Nothing
			End If

			FileName = ""
		End If


		'' 사진
		If PhotoName <>"" Then 
			If FILE_CHECK_EXT(PhotoName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("PhotoName").FileLen Then 
					PhotoName = DextFileUpload("PhotoName",UPLOAD_BASE_PATH & savePathPhoto,0)
				Else
					With Response
					 .Write "<script language='javascript' type='text/javascript'>"
					 .Write "alert('파일의 크기는 50MB 를 넘길수 없습니다.');"
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

			If oldPhotoName <> "" Then
				Set FSO = CreateObject("Scripting.FileSystemObject")
					If (FSO.FileExists(UPLOAD_BASE_PATH & savePathPhoto & oldPhotoName)) Then	' 같은 이름의 파일이 있을 때 삭제
						fso.deletefile(UPLOAD_BASE_PATH & savePathPhoto & oldPhotoName)
					End If
				set FSO = Nothing
			End If
		Else
			PhotoName = oldPhotoName
		End If


		Call Update()
		alertMsg = "수정 되었습니다."
	ElseIf (actType = "DELETE") Then	'글삭제
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
	"DECLARE @UserIdx INT; " &_
	"SET @UserIdx = ? ; " &_

	"INSERT INTO [dbo].[SP_JOB_USER]( " &_
	"	 [Form] " &_
	"	,[Kind] " &_
	"	,[WorkArea] " &_
	"	,[Pay] " &_
	"	,[School] " &_
	"	,[Bigo] " &_
	"	,[File] " &_
	"	,[InData] " &_
	"	,[UserIdx] " &_
	"	,[Ip] " &_
	"	,[Dellfg] " &_
	")VALUES(" &_
	"	 ? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,getDate() " &_
	"	,@UserIdx " &_
	"	,? " &_
	"	,0 " &_
	");" &_

	"DELETE [dbo].[SP_JOB_USER_CAREER] WHERE [UserIdx] = @UserIdx; " &_
	"DELETE [dbo].[SP_JOB_USER_QUALIFY] WHERE [UserIdx] = @UserIdx; " &_

	"DECLARE @CAREER_NAME VARCHAR (max),@CAREER_MONTH VARCHAR (max),@CAREER_POSITION VARCHAR (max) " &_
	"DECLARE @CAREER_T TABLE(NAME VARCHAR(200) , MONTH VARCHAR(100) , POSITION VARCHAR(200) ) " &_
	"SET @CAREER_NAME = ? " &_
	"SET @CAREER_MONTH = ? " &_
	"SET @CAREER_POSITION = ? " &_

	

	"WHILE CHARINDEX(',',@CAREER_NAME)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @CAREER_T(NAME,MONTH,POSITION) VALUES( SUBSTRING(@CAREER_NAME,1,CHARINDEX(',',@CAREER_NAME)-1) , SUBSTRING(@CAREER_MONTH,1,CHARINDEX(',',@CAREER_MONTH)-1) , SUBSTRING(@CAREER_POSITION,1,CHARINDEX(',',@CAREER_POSITION)-1) ) " &_
	"	SET @CAREER_NAME=SUBSTRING(@CAREER_NAME,CHARINDEX(',',@CAREER_NAME)+1,LEN(@CAREER_NAME))  " &_
	"	SET @CAREER_MONTH=SUBSTRING(@CAREER_MONTH,CHARINDEX(',',@CAREER_MONTH)+1,LEN(@CAREER_MONTH))  " &_
	"	SET @CAREER_POSITION=SUBSTRING(@CAREER_POSITION,CHARINDEX(',',@CAREER_POSITION)+1,LEN(@CAREER_POSITION))  " &_
	"END " &_
	"IF CHARINDEX(',',@CAREER_NAME)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @CAREER_T(NAME,MONTH,POSITION) VALUES( @CAREER_NAME , @CAREER_MONTH , @CAREER_POSITION  ) " &_
	"END " &_

	"INSERT INTO [dbo].[SP_JOB_USER_CAREER]( " &_
	"	 [Name] " &_
	"	,[WorkMonth] " &_
	"	,[LastPosition] " &_
	"	,[UserIdx] " &_
	")SELECT " &_
	"	  [NAME]" &_
	"	 ,[MONTH]" &_
	"	 ,[POSITION]" &_
	"	 ,@UserIdx " &_
	"FROM @CAREER_T " &_
	"WHERE [NAME] is not null AND [MONTH] is not null AND [POSITION] is not null " &_
	"AND [NAME] <> '' AND [MONTH] <> '' AND [POSITION] <> '' " &_



	"DECLARE @QUALIFY_NAME VARCHAR (max),@QUALIFY_DATE VARCHAR (max),@QUALIFY_PUBLISH VARCHAR (max) " &_
	"DECLARE @QUALIFY_T TABLE(NAME VARCHAR(200) , DATE VARCHAR(100) , PUBLISH VARCHAR(200) ) " &_
	"SET @QUALIFY_NAME = ? " &_
	"SET @QUALIFY_DATE = ? " &_
	"SET @QUALIFY_PUBLISH = ? " &_

	"WHILE CHARINDEX(',',@QUALIFY_NAME)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @QUALIFY_T(NAME,DATE,PUBLISH) VALUES( SUBSTRING(@QUALIFY_NAME,1,CHARINDEX(',',@QUALIFY_NAME)-1) , SUBSTRING(@QUALIFY_DATE,1,CHARINDEX(',',@QUALIFY_DATE)-1) , SUBSTRING(@QUALIFY_PUBLISH,1,CHARINDEX(',',@QUALIFY_PUBLISH)-1) ) " &_
	"	SET @QUALIFY_NAME=SUBSTRING(@QUALIFY_NAME,CHARINDEX(',',@QUALIFY_NAME)+1,LEN(@QUALIFY_NAME))  " &_
	"	SET @QUALIFY_DATE=SUBSTRING(@QUALIFY_DATE,CHARINDEX(',',@QUALIFY_DATE)+1,LEN(@QUALIFY_DATE))  " &_
	"	SET @QUALIFY_PUBLISH=SUBSTRING(@QUALIFY_PUBLISH,CHARINDEX(',',@QUALIFY_PUBLISH)+1,LEN(@QUALIFY_PUBLISH))  " &_
	"END " &_
	"IF CHARINDEX(',',@QUALIFY_NAME)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @QUALIFY_T(NAME,DATE,PUBLISH) VALUES( @QUALIFY_NAME , @QUALIFY_DATE , @QUALIFY_PUBLISH  ) " &_
	"END " &_

	"INSERT INTO [dbo].[SP_JOB_USER_QUALIFY]( " &_
	"	 [Name] " &_
	"	,[Tdate] " &_
	"	,[Publish] " &_
	"	,[UserIdx] " &_
	")SELECT " &_
	"	  [NAME]" &_
	"	 ,[DATE]" &_
	"	 ,[PUBLISH]" &_
	"	 ,@UserIdx " &_
	"FROM @QUALIFY_T " &_
	"WHERE [NAME] is not null AND [DATE] is not null AND [PUBLISH] is not null " &_
	"AND [NAME] <> '' AND [DATE] <> '' AND [PUBLISH] <> '' " &_

	"UPDATE [dbo].[SP_USER_MEMBER] SET [Photo] = ? WHERE UserIdx = @UserIdx ;"
	
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger      , adParamInput, 0         , session("UserIdx") )
		
		.Parameters.Append .CreateParameter( "@Form"     ,adInteger     , adParamInput, 0         , Form )
		.Parameters.Append .CreateParameter( "@Kind"     ,adInteger     , adParamInput, 0         , Kind )
		.Parameters.Append .CreateParameter( "@WorkArea" ,adVarChar     , adParamInput, 200       , WorkArea )
		.Parameters.Append .CreateParameter( "@Pay"      ,adVarChar     , adParamInput, 200       , Pay )
		.Parameters.Append .CreateParameter( "@School"   ,adVarChar     , adParamInput, 200       , School )
		.Parameters.Append .CreateParameter( "@Bigo"     ,adLongVarChar , adParamInput, 2147483647, Bigo )
		.Parameters.Append .CreateParameter( "@File"     ,adVarChar     , adParamInput, 200       , FileName )

		.Parameters.Append .CreateParameter( "@g_uip"    ,adVarChar     , adParamInput, 20        , g_uip )

		.Parameters.Append .CreateParameter( "@CAREER_NAME"     ,adLongVarChar , adParamInput, 2147483647 , careerName )
		.Parameters.Append .CreateParameter( "@CAREER_MONTH"    ,adLongVarChar , adParamInput, 2147483647 , careerMonth )
		.Parameters.Append .CreateParameter( "@CAREER_POSITION" ,adLongVarChar , adParamInput, 2147483647 , careerPosition )

		.Parameters.Append .CreateParameter( "@QUALIFY_NAME"    ,adLongVarChar , adParamInput, 2147483647 , qualifyName )
		.Parameters.Append .CreateParameter( "@QUALIFY_DATE"    ,adLongVarChar , adParamInput, 2147483647 , qualifyDate )
		.Parameters.Append .CreateParameter( "@QUALIFY_PUBLISH" ,adLongVarChar , adParamInput, 2147483647 , qualifyPublish )

		.Parameters.Append .CreateParameter( "@Photo"           ,adVarChar     , adParamInput, 200       , PhotoName )


		.Execute
	End with
	call cmdclose()
End Sub

'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @UserIdx INT , @Idx INT ; " &_
	"SET @UserIdx = ? ; " &_
	"SET @Idx     = ? ; " &_
	
	"UPDATE [dbo].[SP_JOB_USER] SET " &_
	"	 [Form]     = ? " &_
	"	,[Kind]     = ? " &_
	"	,[WorkArea] = ? " &_
	"	,[Pay]      = ? " &_
	"	,[School]   = ? " &_
	"	,[Bigo]     = ? " &_
	"	,[File]     = ? " &_
	"WHERE [Idx] = @Idx AND [UserIdx] = @UserIdx ; " &_

	"DELETE [dbo].[SP_JOB_USER_CAREER] WHERE [UserIdx] = @UserIdx; " &_
	"DELETE [dbo].[SP_JOB_USER_QUALIFY] WHERE [UserIdx] = @UserIdx; " &_

	"DECLARE @CAREER_NAME VARCHAR (max),@CAREER_MONTH VARCHAR (max),@CAREER_POSITION VARCHAR (max) " &_
	"DECLARE @CAREER_T TABLE(NAME VARCHAR(200) , MONTH VARCHAR(100) , POSITION VARCHAR(200) ) " &_
	"SET @CAREER_NAME = ? " &_
	"SET @CAREER_MONTH = ? " &_
	"SET @CAREER_POSITION = ? " &_

	

	"WHILE CHARINDEX(',',@CAREER_NAME)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @CAREER_T(NAME,MONTH,POSITION) VALUES( SUBSTRING(@CAREER_NAME,1,CHARINDEX(',',@CAREER_NAME)-1) , SUBSTRING(@CAREER_MONTH,1,CHARINDEX(',',@CAREER_MONTH)-1) , SUBSTRING(@CAREER_POSITION,1,CHARINDEX(',',@CAREER_POSITION)-1) ) " &_
	"	SET @CAREER_NAME=SUBSTRING(@CAREER_NAME,CHARINDEX(',',@CAREER_NAME)+1,LEN(@CAREER_NAME))  " &_
	"	SET @CAREER_MONTH=SUBSTRING(@CAREER_MONTH,CHARINDEX(',',@CAREER_MONTH)+1,LEN(@CAREER_MONTH))  " &_
	"	SET @CAREER_POSITION=SUBSTRING(@CAREER_POSITION,CHARINDEX(',',@CAREER_POSITION)+1,LEN(@CAREER_POSITION))  " &_
	"END " &_
	"IF CHARINDEX(',',@CAREER_NAME)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @CAREER_T(NAME,MONTH,POSITION) VALUES( @CAREER_NAME , @CAREER_MONTH , @CAREER_POSITION  ) " &_
	"END " &_

	"INSERT INTO [dbo].[SP_JOB_USER_CAREER]( " &_
	"	 [Name] " &_
	"	,[WorkMonth] " &_
	"	,[LastPosition] " &_
	"	,[UserIdx] " &_
	")SELECT " &_
	"	  [NAME]" &_
	"	 ,[MONTH]" &_
	"	 ,[POSITION]" &_
	"	 ,@UserIdx " &_
	"FROM @CAREER_T " &_
	"WHERE [NAME] is not null AND [MONTH] is not null AND [POSITION] is not null " &_
	"AND [NAME] <> '' AND [MONTH] <> '' AND [POSITION] <> '' " &_


	"DECLARE @QUALIFY_NAME VARCHAR (max),@QUALIFY_DATE VARCHAR (max),@QUALIFY_PUBLISH VARCHAR (max) " &_
	"DECLARE @QUALIFY_T TABLE(NAME VARCHAR(200) , DATE VARCHAR(100) , PUBLISH VARCHAR(200) ) " &_
	"SET @QUALIFY_NAME = ? " &_
	"SET @QUALIFY_DATE = ? " &_
	"SET @QUALIFY_PUBLISH = ? " &_

	"WHILE CHARINDEX(',',@QUALIFY_NAME)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @QUALIFY_T(NAME,DATE,PUBLISH) VALUES( SUBSTRING(@QUALIFY_NAME,1,CHARINDEX(',',@QUALIFY_NAME)-1) , SUBSTRING(@QUALIFY_DATE,1,CHARINDEX(',',@QUALIFY_DATE)-1) , SUBSTRING(@QUALIFY_PUBLISH,1,CHARINDEX(',',@QUALIFY_PUBLISH)-1) ) " &_
	"	SET @QUALIFY_NAME=SUBSTRING(@QUALIFY_NAME,CHARINDEX(',',@QUALIFY_NAME)+1,LEN(@QUALIFY_NAME))  " &_
	"	SET @QUALIFY_DATE=SUBSTRING(@QUALIFY_DATE,CHARINDEX(',',@QUALIFY_DATE)+1,LEN(@QUALIFY_DATE))  " &_
	"	SET @QUALIFY_PUBLISH=SUBSTRING(@QUALIFY_PUBLISH,CHARINDEX(',',@QUALIFY_PUBLISH)+1,LEN(@QUALIFY_PUBLISH))  " &_
	"END " &_
	"IF CHARINDEX(',',@QUALIFY_NAME)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @QUALIFY_T(NAME,DATE,PUBLISH) VALUES( @QUALIFY_NAME , @QUALIFY_DATE , @QUALIFY_PUBLISH  ) " &_
	"END " &_

	"INSERT INTO [dbo].[SP_JOB_USER_QUALIFY]( " &_
	"	 [Name] " &_
	"	,[Tdate] " &_
	"	,[Publish] " &_
	"	,[UserIdx] " &_
	")SELECT " &_
	"	  [NAME]" &_
	"	 ,[DATE]" &_
	"	 ,[PUBLISH]" &_
	"	 ,@UserIdx " &_
	"FROM @QUALIFY_T " &_
	"WHERE [NAME] is not null AND [DATE] is not null AND [PUBLISH] is not null " &_
	"AND [NAME] <> '' AND [DATE] <> '' AND [PUBLISH] <> '' " &_

	"UPDATE [dbo].[SP_USER_MEMBER] SET [Photo] = ? WHERE UserIdx = @UserIdx ;"




	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx"  ,adInteger      , adParamInput, 0         , session("UserIdx") )
		.Parameters.Append .CreateParameter( "@Idx"      ,adInteger      , adParamInput, 0         , Idx )

		.Parameters.Append .CreateParameter( "@Form"     ,adInteger     , adParamInput, 0         , Form )
		.Parameters.Append .CreateParameter( "@Kind"     ,adInteger     , adParamInput, 0         , Kind )
		.Parameters.Append .CreateParameter( "@WorkArea" ,adVarChar     , adParamInput, 200       , WorkArea )
		.Parameters.Append .CreateParameter( "@Pay"      ,adVarChar     , adParamInput, 200       , Pay )
		.Parameters.Append .CreateParameter( "@School"   ,adVarChar     , adParamInput, 200       , School )
		.Parameters.Append .CreateParameter( "@Bigo"     ,adLongVarChar , adParamInput, 2147483647, Bigo )
		.Parameters.Append .CreateParameter( "@File"     ,adVarChar     , adParamInput, 200       , FileName )

		.Parameters.Append .CreateParameter( "@CAREER_NAME"     ,adLongVarChar , adParamInput, 2147483647 , careerName )
		.Parameters.Append .CreateParameter( "@CAREER_MONTH"    ,adLongVarChar , adParamInput, 2147483647 , careerMonth )
		.Parameters.Append .CreateParameter( "@CAREER_POSITION" ,adLongVarChar , adParamInput, 2147483647 , careerPosition )

		.Parameters.Append .CreateParameter( "@QUALIFY_NAME"    ,adLongVarChar , adParamInput, 2147483647 , qualifyName )
		.Parameters.Append .CreateParameter( "@QUALIFY_DATE"    ,adLongVarChar , adParamInput, 2147483647 , qualifyDate )
		.Parameters.Append .CreateParameter( "@QUALIFY_PUBLISH" ,adLongVarChar , adParamInput, 2147483647 , qualifyPublish )

		.Parameters.Append .CreateParameter( "@Photo"           ,adVarChar     , adParamInput, 200       , PhotoName )
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
	
	
	"UPDATE [dbo].[SP_JOB_USER] SET " &_
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
	top.location.href = "fJobL.asp?<%=PageParams%>";
</script>
</html>