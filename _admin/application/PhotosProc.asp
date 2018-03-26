<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim savePath   : savePath = "\appMember/" '첨부 저장경로.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath
UPLOAD__FORM.MaxFileLen		= 10 * 1024 * 1024 '10메가

Dim actType      : actType   = UPLOAD__FORM("actType")
Dim UserIdx      : UserIdx   = UPLOAD__FORM("UserIdx")
Dim GoUrl        : GoUrl     = UPLOAD__FORM("GoUrl")
Dim PhotoName    : PhotoName = UPLOAD__FORM("PhotoName")
Dim oldPhotoName : oldPhotoName = UPLOAD__FORM("oldPhotoName")

If UserIdx = "" Then 
	With Response
	 .Write "<script language='javascript' type='text/javascript'>"
	 .Write "alert('회원정보가 누락되었습니다. 관리자에게 문의바랍니다.');"
	 .Write "history.go(-1);"
	 .Write "</script>"
	 .End
	End With
End If

Call Expires()
Call dbopen()

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
	
Call dbclose()



'수정
Sub Update()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @UserIdx INT,@Photo VARCHAR (200);" &_
	"SET @UserIdx = ? ; " &_
	"SET @Photo   = ? ; " &_

	"UPDATE [dbo].[SP_USER_MEMBER] SET [Photo] = @Photo WHERE UserIdx = @UserIdx ;"
	
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput, 0   , UserIdx )
		.Parameters.Append .CreateParameter( "@File"    ,adVarChar , adParamInput, 200 , PhotoName )
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
	top.location.href = "<%=GoUrl%>";
</script>
</html>