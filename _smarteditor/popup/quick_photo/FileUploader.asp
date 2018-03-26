<!-- #include file = "../../../_lib/common.asp" -->
<!-- #include file = "../../../_lib/uploadUtil.asp" -->
<%
Dim savePath : savePath = "\SmtEdit/" '첨부 저장경로.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath

Dim upload_file   : upload_file   = UPLOAD__FORM("uploadInputBox")
Dim callback_func : callback_func = UPLOAD__FORM("callback_func")

If upload_file <> "" Then
	upload_file = DextFileUpload("uploadInputBox",UPLOAD_BASE_PATH & savePath,0)
End If

Dim url : url = g_host & FRONT_ROOT_DIR & "upload/SmtEdit/" & upload_file

Response.redirect "callback.html?callback_func=" & callback_func & "&bNewLine=true&sFileName=" & upload_file & "&sFileURL=" & url
%>