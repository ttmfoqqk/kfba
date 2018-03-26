<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
If Session("UserIdx") <> "" Then 
	Response.redirect "../mypage/"
End If
Dim UserName   : UserName = RequestSet("UserName","POST","")
Dim UserBirth1 : UserBirth1 = RequestSet("UserBirth1","POST","")
Dim UserBirth2 : UserBirth2 = RequestSet("UserBirth2","POST","")
Dim UserBirth3 : UserBirth3 = RequestSet("UserBirth3","POST","")
Dim UserPhone1 : UserPhone1 = RequestSet("UserPhone1","POST","")
Dim UserPhone2 : UserPhone2 = RequestSet("UserPhone2","POST","")
Dim UserPhone3 : UserPhone3 = RequestSet("UserPhone3","POST","")

Dim UserEmail1 : UserEmail1 = RequestSet("UserEmail1","POST","")
Dim UserEmail2 : UserEmail2 = RequestSet("UserEmail2","POST","")

Dim sMode      : sMode      = RequestSet("sMode","POST","")

Dim ResultMsg  : ResultMsg  = "�Է��Ͻ� ������ ��ġ�ϴ� ���̵� �����ϴ�.<br>��Ȯ�� ������ Ȯ�� �� �ٽ� �Է� ��Ź �帳�ϴ�."

Call Expires()
Call dbopen()
	If sMode = "phone" Then 
		Call getDataPhone()
	Else
		Call getDataEmail()
	End If
Call dbclose()

If FI_UserId <> "" Then 
	ResultMsg = "������ ���̵�� <b>" & FI_UserId & "</b> �Դϴ�."
End If 

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "member/fIdResult.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("ResultMsg", ResultMsg ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = Nothing


Sub getDataPhone()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT top 1 [UserId] " &_
	"FROM [dbo].[SP_USER_MEMBER] " &_
	"WHERE [UserName] = ? " &_
	"AND [UserBirth] = ? " &_
	"AND [UserHphone1] = ? " &_
	"AND [UserHphone2] = ? " &_
	"AND [UserHphone3] = ? " &_
	"AND [UserDelFg] = 0 ORDER BY [UserIdx] DESC"
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserName"   ,adVarChar , adParamInput, 50 , UserName )
		.Parameters.Append .CreateParameter( "@UserBirth"  ,adVarChar , adParamInput, 8  , UserBirth1 & UserBirth2 & UserBirth3 )
		.Parameters.Append .CreateParameter( "@UserPhone1" ,adVarChar , adParamInput, 4  , UserPhone1 )
		.Parameters.Append .CreateParameter( "@UserPhone2" ,adVarChar , adParamInput, 4  , UserPhone2 )
		.Parameters.Append .CreateParameter( "@UserPhone3" ,adVarChar , adParamInput, 4  , UserPhone3 )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub

Sub getDataEmail()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT [UserId] " &_
	"FROM [dbo].[SP_USER_MEMBER] " &_
	"WHERE [UserName] = ? " &_
	"AND [UserBirth] = ? " &_
	"AND [UserEmail] = ? " &_
	"AND [UserDelFg] = 0 ORDER BY [UserIdx] DESC"
	
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserName"  ,adVarChar , adParamInput, 50   , UserName )
		.Parameters.Append .CreateParameter( "@UserBirth" ,adVarChar , adParamInput, 8    , UserBirth1 & UserBirth2 & UserBirth3 )
		.Parameters.Append .CreateParameter( "@UserEmail" ,adVarChar , adParamInput, 1000 , UserEmail1 &"@"& UserEmail2 )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
%>