<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim savePath   : savePath = "\appMember/" '첨부 저장경로.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath
UPLOAD__FORM.MaxFileLen		= 10 * 1024 * 1024 '10메가

Dim alertMsg       : alertMsg = ""
Dim actType        : actType        = UPLOAD__FORM("actType")

Dim Idx            : Idx            = UPLOAD__FORM("Idx")
Dim ProgramIdx     : ProgramIdx     = IIF( UPLOAD__FORM("ProgramIdx")=""  , 0 , UPLOAD__FORM("ProgramIdx") )
Dim ProgramKind    : ProgramKind    = IIF( UPLOAD__FORM("ProgramKind")="" , 0 , UPLOAD__FORM("ProgramKind") )
Dim PhotoName      : PhotoName      = UPLOAD__FORM("PhotoName")
Dim oldPhotoName   : oldPhotoName   = UPLOAD__FORM("oldPhotoName")

Dim FileName       : FileName       = UPLOAD__FORM("FileName")
Dim oldFileName    : oldFileName    = UPLOAD__FORM("oldFileName")
Dim DellFileFg     : DellFileFg     = UPLOAD__FORM("DellFileFg")

Dim CompanyName    : CompanyName    = UPLOAD__FORM("CompanyName")
Dim WorkTime       : WorkTime       = UPLOAD__FORM("WorkTime")
Dim WorkMonth      : WorkMonth      = UPLOAD__FORM("WorkMonth")
Dim LastPosition   : LastPosition   = UPLOAD__FORM("LastPosition")

Dim pageNo         : pageNo         = IIF(UPLOAD__FORM("pageNo")="",1,UPLOAD__FORM("pageNo"))
Dim pagePosition   : pagePosition   = IIF(UPLOAD__FORM("pagePosition")="","",UPLOAD__FORM("pagePosition"))

Dim GoUrl : GoUrl = "../judge/"
Dim PageParams
PageParams = "pageNo=" & pageNo

Call Expires()
Call dbopen()

	if (alertMsg <> "")	then
		actType	= ""
	Elseif (actType = "INSERT") Then	'글작성
		If( ProgramIdx = "" Or ProgramKind = "" ) Then 
			Call msgbox("잘못된 경로 입니다.", true)
		End If
		Call Check()

		If FI_CntDuplicate > 0 Then 
			Call msgbox("이미 접수된 자격종목 입니다.", true)
		End If 

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
		
		'' 사진
		If PhotoName <>"" Then 
			If FILE_CHECK_EXT(PhotoName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("PhotoName").FileLen Then 
					PhotoName = DextFileUpload("PhotoName",UPLOAD_BASE_PATH & savePath,0)
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
		Else 
			PhotoName = oldPhotoName
		End If

		Call Insert()
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
		
		'' 사진
		If PhotoName <>"" Then 
			If FILE_CHECK_EXT(PhotoName) = True Then
				If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("PhotoName").FileLen Then 
					PhotoName = DextFileUpload("PhotoName",UPLOAD_BASE_PATH & savePath,0)
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

			If oldPhotoName <> "" Then
				Set FSO = CreateObject("Scripting.FileSystemObject")
					If (FSO.FileExists(UPLOAD_BASE_PATH & savePath & oldPhotoName)) Then	' 같은 이름의 파일이 있을 때 삭제
						fso.deletefile(UPLOAD_BASE_PATH & savePath & oldPhotoName)
					End If
				set FSO = Nothing
			End If
		Else
			PhotoName = oldPhotoName
		End If


		Call Update()
		alertMsg = "수정 되었습니다."

		If pagePosition = "mypage" Then 
			GoUrl = "../mypage/judge.asp"
		End If
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
	"DECLARE @UserIdx INT , @ProgramIdx INT , @ProgramKind INT ; " &_
	"SET @UserIdx = ? ; " &_
	"SET @ProgramIdx  = ? ; " &_
	"SET @ProgramKind = ? ; " &_

	"INSERT INTO [dbo].[SP_PROGRAM_JUDGE_APP]( " &_
	"	 [ProgramIdx] " &_
	"	,[UserIdx] " &_
	"	,[State] " &_
	"	,[InData] " &_
	"	,[Ip] " &_
	"	,[Dellfg] " &_
	"	,[FileName] " &_
	"	,[ProgramKind] " &_
	")VALUES(" &_
	"	 @ProgramIdx " &_
	"	,@UserIdx " &_
	"	,1 " &_
	"	,getDate() " &_
	"	,? " &_
	"	,0 " &_
	"	,? " &_
	"	,@ProgramKind " &_
	");" &_

	"DELETE [dbo].[SP_PROGRAM_JUDGE_APP_CAREER] WHERE [UserIdx] = @UserIdx; " &_

	"DECLARE @CompanyName VARCHAR (max),@WorkTime VARCHAR (max),@WorkMonth VARCHAR (max) ,@LastPosition VARCHAR (max) " &_
	"DECLARE @T TABLE(CompanyName VARCHAR(200) , WorkTime VARCHAR(100) , WorkMonth VARCHAR(200) , LastPosition VARCHAR(200) ) " &_
	"SET @CompanyName = ? " &_
	"SET @WorkTime = ? " &_
	"SET @WorkMonth = ? " &_
	"SET @LastPosition = ? " &_

	

	"WHILE CHARINDEX(',',@CompanyName)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(CompanyName,WorkTime,WorkMonth,LastPosition) VALUES( SUBSTRING(@CompanyName,1,CHARINDEX(',',@CompanyName)-1) , SUBSTRING(@WorkTime,1,CHARINDEX(',',@WorkTime)-1) , SUBSTRING(@WorkTime,1,CHARINDEX(',',@WorkTime)-1) , SUBSTRING(@LastPosition,1,CHARINDEX(',',@LastPosition)-1) ) " &_
	"	SET @CompanyName=SUBSTRING(@CompanyName,CHARINDEX(',',@CompanyName)+1,LEN(@CompanyName))  " &_
	"	SET @WorkTime=SUBSTRING(@WorkTime,CHARINDEX(',',@WorkTime)+1,LEN(@WorkTime))  " &_
	"	SET @WorkMonth=SUBSTRING(@WorkMonth,CHARINDEX(',',@WorkMonth)+1,LEN(@WorkMonth))  " &_
	"	SET @LastPosition=SUBSTRING(@LastPosition,CHARINDEX(',',@LastPosition)+1,LEN(@LastPosition))  " &_
	"END " &_
	"IF CHARINDEX(',',@CompanyName)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(CompanyName,WorkTime,WorkMonth,LastPosition) VALUES( @CompanyName , @WorkTime , @WorkMonth ,@LastPosition  ) " &_
	"END " &_

	"INSERT INTO [dbo].[SP_PROGRAM_JUDGE_APP_CAREER]( " &_
	"	 [CompanyName] " &_
	"	,[WorkTime] " &_
	"	,[WorkMonth] " &_
	"	,[LastPosition] " &_
	"	,[UserIdx] " &_
	")SELECT " &_
	"	  LTRIM(  RTRIM( [CompanyName] ) )" &_
	"	 ,LTRIM(  RTRIM( [WorkTime] ) )" &_
	"	 ,LTRIM(  RTRIM( [WorkMonth] ) )" &_
	"	 ,LTRIM(  RTRIM( [LastPosition] ) )" &_
	"	 ,@UserIdx " &_
	"FROM @T " &_
	"WHERE [CompanyName] is not null AND [WorkTime] is not null AND [WorkMonth] is not null AND [LastPosition] is not null " &_
	"AND [CompanyName] <> '' AND [WorkTime] <> '' AND [WorkMonth] <> '' AND [LastPosition] <> '' " &_


	"UPDATE [dbo].[SP_USER_MEMBER] SET [Photo] = ? WHERE UserIdx = @UserIdx ;"
	
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx"      ,adInteger     , adParamInput, 0          , session("UserIdx") )
		.Parameters.Append .CreateParameter( "@ProgramIdx"   ,adInteger     , adParamInput, 0          , ProgramIdx )
		.Parameters.Append .CreateParameter( "@ProgramKind"  ,adInteger     , adParamInput, 0          , ProgramKind )

		.Parameters.Append .CreateParameter( "@g_uip"        ,adVarChar     , adParamInput, 20         , g_uip )
		.Parameters.Append .CreateParameter( "@File"         ,adVarChar     , adParamInput, 200        , FileName )

		.Parameters.Append .CreateParameter( "@CompanyName"  ,adLongVarChar , adParamInput, 2147483647 , CompanyName )
		.Parameters.Append .CreateParameter( "@WorkTime"     ,adLongVarChar , adParamInput, 2147483647 , WorkTime )
		.Parameters.Append .CreateParameter( "@WorkMonth"    ,adLongVarChar , adParamInput, 2147483647 , WorkMonth )
		.Parameters.Append .CreateParameter( "@LastPosition" ,adLongVarChar , adParamInput, 2147483647 , LastPosition )


		.Parameters.Append .CreateParameter( "@Photo"        ,adVarChar     , adParamInput, 200        , PhotoName )


		.Execute
	End with
	call cmdclose()
End Sub

'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @UserIdx INT;" &_
	"SET @UserIdx = ? ; " &_

	"DELETE [dbo].[SP_PROGRAM_JUDGE_APP_CAREER] WHERE [UserIdx] = @UserIdx; " &_

	"DECLARE @CompanyName VARCHAR (max),@WorkTime VARCHAR (max),@WorkMonth VARCHAR (max) ,@LastPosition VARCHAR (max) " &_
	"DECLARE @T TABLE(CompanyName VARCHAR(200) , WorkTime VARCHAR(100) , WorkMonth VARCHAR(200) , LastPosition VARCHAR(200) ) " &_
	"SET @CompanyName = ? " &_
	"SET @WorkTime = ? " &_
	"SET @WorkMonth = ? " &_
	"SET @LastPosition = ? " &_


	"WHILE CHARINDEX(',',@CompanyName)<>0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(CompanyName,WorkTime,WorkMonth,LastPosition) VALUES( SUBSTRING(@CompanyName,1,CHARINDEX(',',@CompanyName)-1) , SUBSTRING(@WorkTime,1,CHARINDEX(',',@WorkTime)-1) , SUBSTRING(@WorkTime,1,CHARINDEX(',',@WorkTime)-1) , SUBSTRING(@LastPosition,1,CHARINDEX(',',@LastPosition)-1) ) " &_
	"	SET @CompanyName=SUBSTRING(@CompanyName,CHARINDEX(',',@CompanyName)+1,LEN(@CompanyName))  " &_
	"	SET @WorkTime=SUBSTRING(@WorkTime,CHARINDEX(',',@WorkTime)+1,LEN(@WorkTime))  " &_
	"	SET @WorkMonth=SUBSTRING(@WorkMonth,CHARINDEX(',',@WorkMonth)+1,LEN(@WorkMonth))  " &_
	"	SET @LastPosition=SUBSTRING(@LastPosition,CHARINDEX(',',@LastPosition)+1,LEN(@LastPosition))  " &_
	"END " &_
	"IF CHARINDEX(',',@CompanyName)=0 " &_
	"BEGIN " &_
	"	INSERT INTO @T(CompanyName,WorkTime,WorkMonth,LastPosition) VALUES( @CompanyName , @WorkTime , @WorkMonth ,@LastPosition  ) " &_
	"END " &_

	"INSERT INTO [dbo].[SP_PROGRAM_JUDGE_APP_CAREER]( " &_
	"	 [CompanyName] " &_
	"	,[WorkTime] " &_
	"	,[WorkMonth] " &_
	"	,[LastPosition] " &_
	"	,[UserIdx] " &_
	")SELECT " &_
	"	  LTRIM(  RTRIM( [CompanyName] ) )" &_
	"	 ,LTRIM(  RTRIM( [WorkTime] ) )" &_
	"	 ,LTRIM(  RTRIM( [WorkMonth] ) )" &_
	"	 ,LTRIM(  RTRIM( [LastPosition] ) )" &_
	"	 ,@UserIdx " &_
	"FROM @T " &_
	"WHERE [CompanyName] is not null AND [WorkTime] is not null AND [WorkMonth] is not null AND [LastPosition] is not null " &_
	"AND [CompanyName] <> '' AND [WorkTime] <> '' AND [WorkMonth] <> '' AND [LastPosition] <> '' " &_


	"UPDATE [dbo].[SP_USER_MEMBER] SET [Photo] = ? WHERE UserIdx = @UserIdx ;" &_

	"UPDATE [dbo].[SP_PROGRAM_JUDGE_APP] SET [FileName] = ? WHERE [Idx] = ? ;"
	
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx"      ,adInteger     , adParamInput, 0          , session("UserIdx") )

		.Parameters.Append .CreateParameter( "@CompanyName"  ,adLongVarChar , adParamInput, 2147483647 , CompanyName )
		.Parameters.Append .CreateParameter( "@WorkTime"     ,adLongVarChar , adParamInput, 2147483647 , WorkTime )
		.Parameters.Append .CreateParameter( "@WorkMonth"    ,adLongVarChar , adParamInput, 2147483647 , WorkMonth )
		.Parameters.Append .CreateParameter( "@LastPosition" ,adLongVarChar , adParamInput, 2147483647 , LastPosition )

		.Parameters.Append .CreateParameter( "@Photo"        ,adVarChar     , adParamInput, 200        , PhotoName )

		.Parameters.Append .CreateParameter( "@File"         ,adVarChar     , adParamInput, 200        , FileName )
		.Parameters.Append .CreateParameter( "@Idx"          ,adInteger     , adParamInput, 0          , Idx )
		.Execute
	End with
	call cmdclose()
End Sub


Sub Check()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @ProgramIdx INT,@ProgramKind INT, @UserIdx INT ,@CntDuplicate INT; " &_
	"SET @ProgramIdx  = ?; " &_
	"SET @ProgramKind = ?; " &_
	"SET @UserIdx     = ?; " &_
	
	"SET @CntDuplicate = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM_JUDGE_APP] WHERE [ProgramIdx] = @ProgramIdx AND [ProgramKind] = @ProgramKind AND [UserIdx] = @UserIdx AND [Dellfg] = 0 AND [State] != 2 )  " &_

	"SELECT @CntDuplicate AS [CntDuplicate] "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@ProgramIdx"  ,adInteger , adParamInput, 0 , ProgramIdx )
		.Parameters.Append .CreateParameter( "@ProgramKind" ,adInteger , adParamInput, 0 , ProgramKind )
		.Parameters.Append .CreateParameter( "@UserIdx"     ,adInteger , adParamInput, 0 , session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()
	'프로그램정보
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN">
<html>
<head>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
</head>
<script language=javascript>
	if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
	top.location.href = "<%=GoUrl%>?<%=PageParams%>";
</script>
</html>