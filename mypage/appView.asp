<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%

Dim rows     : rows     = 10
Dim pageNo   : pageNo   = RequestSet("pageNo"   ,"GET",1 )
Dim Idx      : Idx      = RequestSet("Idx"   ,"GET",0 )
Dim PageParams
PageParams = "pageNo=" & pageNo & "&Idx=" & Idx

checkLogin( g_host & g_url &"?"&PageParams )

Call Expires()
Call dbopen()
	Call getView()
	Call CheckApplicationCnt()
Call dbclose()

Dim StateTxt

If FV_State = "0" Then 
	StateTxt = "�����Ϸ�"
ElseIf FV_State = "1" Then 
	StateTxt = "<font color=""#11179a"">�Աݴ��</font>"
ElseIf FV_State = "2" Then 
	StateTxt = "<font color=""#9a1134"">�������</font>"
ElseIf FV_State = "3" Then 
	StateTxt = "<font color=""#9a1134"">���հ�</font>"
ElseIf FV_State = "4" Then 
	StateTxt = "<font color=""#9a1134"">������(���հ�)</font>"
ElseIf FV_State = "10" Then 
	StateTxt = "<font color=""#11179a"">�հ�</font>"
End If

Dim PayModeTxt
If FV_PayMode = "SC0010" Then 
	PayModeTxt = "ī�����"
ElseIf FV_PayMode = "SC0030" Then 
	PayModeTxt = "�ǽð� ������ü"
ElseIf FV_PayMode = "SC0060" Then 
	PayModeTxt = "�ڵ�������"
ElseIf FV_PayMode = "SC0040" Then 
	PayModeTxt = "��������Ա�"
End If

PrograName = FV_ProgramName

If FV_Class = "1" Then
	PrograName = PrograName & " 1��"
ElseIf FV_Class = "2" Then
	PrograName = PrograName & " 2��"
ElseIf FV_Class = "3" Then
	PrograName = PrograName & " 3��"
End If

If FV_Kind = "1" Then
	PrograName = PrograName & " [�ʱ�]"
ElseIf FV_Kind = "2" Then
	PrograName = PrograName & " [�Ǳ�]"
End If

PhotoExt = FILE_CHECK_EXT_RETURN( FV_Map )
If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
	' ����Ʈ ���� ���� 435
	MapImages = img_resize( "/upload/programsArea/", FV_Map ,435,600)
Else
	MapImages= "<a href=""_lib/dowload.asp?pach=/upload/programsArea/&file="&FV_Map&""">"& FV_Map &"</a>"
End If

PhotoExt = FILE_CHECK_EXT_RETURN( FI_Photo )
If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
	checkSize = Split(imgFileSizeChk(USER_PHOTO_PATH & FI_Photo),"/")
	if checkSize(0) > 0 or checkSize(1) > 0 then 
		UserPhotos = img_resize(USER_PHOTO_PATH, FI_Photo ,150,200)
	else
		UserPhotos= "<a href="""& DOWNLOAD_USER_PHOTO_PATH &  FI_Photo &""">"& FI_Photo &"</a>"
	end if
Else
	UserPhotos= "<a href="""& DOWNLOAD_USER_PHOTO_PATH &  FI_Photo &""">"& FI_Photo &"</a>"
End If

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "mypage/appView.html" ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")
call ntpl.setBlock("MAIN", array("MYJUDGE_LIST"))
'���� �ɻ�������� �޴�
If LEFT_JUDGE_MENU_CNT > 0 Then 
	ntpl.tplParseBlock("MYJUDGE_LIST")
Else
	ntpl.tplBlockDel("MYJUDGE_LIST")
End If 

ntpl.tplAssign array(   _
	 array("imgDir"     , TPL_DIR_IMAGES ) _
	,array("pageNo"     , pageNo ) _
	,array("PageParams" , PageParams ) _

	,array("ProgramName" , PrograName ) _
	,array("OnData"      , FV_OnData ) _
	,array("State"       , StateTxt ) _
	,array("Pay"         , FormatNumber(FV_Pay,0) & " ��" ) _
	,array("PayModeTxt"  , PayModeTxt ) _
	,array("Snumber"     , IIF( FV_Snumber="" , "�����Ϸ� �� ���˹�ȣ�� �ο��˴ϴ�." , FV_Snumber ) ) _
	,array("AreaName"    , FV_Name ) _
	,array("AreaAddr"    , FV_Addr ) _
	,array("AreaTel"     , FV_Tel ) _
	,array("AreaInfo"    , FV_Info ) _
	,array("AreaWebUrl"  , IIF( FV_WebAddr="","&nbsp;", "<a href="""&Replace(FV_WebAddr,"http://","")&""" target=""_blank"">"&FV_WebAddr&"</a> " ) ) _
	,array("AreaMap"     , MapImages ) _

	,array("Photo"       , UserPhotos ) _
	,array("UserName"    , FI_UserName ) _
	,array("UserBirth"   , Mid(FI_UserBirth,1,4) &"."& Mid(FI_UserBirth,5,2) &"."& Mid(FI_UserBirth,7,2) ) _
	,array("oldPhotoName", FI_Photo ) _
	,array("GoUrl"       , g_host & g_url &"?"&PageParams ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = Nothing






Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Idx INT;" &_
	"DECLARE @UserIdx INT;" &_
	"SET @Idx     = ?; " &_
	"SET @UserIdx = ?; " &_

	"SELECT " &_
	"	 ( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = B.[CodeIdx] ) AS [ProgramName] "&_
	"	,B.[OnData] " &_
	"	,A.[State] " &_
	"	,B.[Pay] " &_
	"	,B.[Kind] " &_
	"	,B.[Class] " &_
	"	,A.[PayMode] " &_
	"	,A.[Snumber] " &_
	"	,C.[Name] " &_
	"	,C.[Addr] " &_
	"	,C.[Tel] " &_
	"	,C.[Info] " &_
	"	,C.[WebAddr] " &_
	"	,C.[Map] " &_
	"FROM [dbo].[SP_PROGRAM_APP] A " &_
	"INNER JOIN [dbo].[SP_PROGRAM] B ON(A.[ProgramIdx] = B.[Idx])" &_
	"INNER JOIN [dbo].[SP_PROGRAM_AREA] C ON(A.[AreaIdx] = C.[Idx])" &_
	"WHERE A.[UserIdx] = @UserIdx AND A.[Idx] = @Idx " &_

	"SELECT " &_
	"	 [UserName]" &_
	"	,[UserBirth]" &_
	"	,[Photo]" &_
	"	,[LastName]" &_
	"	,[FirstName]" &_
	"FROM [dbo].[SP_USER_MEMBER] WHERE [UserIdx] = @UserIdx "



	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"     ,adInteger , adParamInput , 0 , Idx )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")

	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "FI")
	objRs.close	: Set objRs = Nothing

	Set objRs = Nothing
End Sub
%>