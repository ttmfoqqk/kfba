<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%

'56 : Ŀ�ǹٸ���Ÿ
'57 : Ĭ�������ֻ�
'58 : �ͼַ�����Ʈ
'59 : ���μҹɸ���
'60 : �����ְ�����
'61 : �ܽİ濵������
'62 : �����������

Dim arrSubData
Dim cntSubData  : cntSubData  = -1
Dim applicationKey : applicationKey = RequestSet("applicationKey","GET",56)
Dim tabOnOff1 : tabOnOff1 = "_off"
Dim tabOnOff2 : tabOnOff2 = "_off"
Dim tabOnOff3 : tabOnOff3 = "_on"


Dim PageParams : PageParams = "applicationKey=" & applicationKey
checkLogin( g_host & g_url &"?"&PageParams )

Dim actType      : actType     = "INSERT"
Dim importFile   : importFile  = "write.html"
Dim ProgramKind  : ProgramKind = "<option value="""">����</option><option value=""1"">�ʱ�</option><option value=""2"">�Ǳ�</option>"
Dim ProgramClass : ProgramClass = "<option value="""">����</option><option value=""1"">1��</option><option value=""2"">2��</option>"

Dim programName

Select Case applicationKey
    Case 56
        programName = "Ŀ�ǹٸ���Ÿ"
		programTitleImg = "01"
		ProgramClass = ProgramClass & "<option value=""3"">3��</option>"
    Case 57
        programName = "Ĭ�������ֻ�"
		programTitleImg = "02"
		importFile      = "write_noClass.html"
    Case 58
        programName = "�ͼַ�����Ʈ"
		programTitleImg = "03"
	Case 59
        programName = "���μҹɸ���"
		programTitleImg = "04"
	Case 60
        programName = "�󶼾�Ʈ ������"
		programTitleImg = "05"
		importFile = "write_noClassKind.html"
	Case 61
        programName = "�ܽİ濵������"
		programTitleImg = "06"
		importFile      = "write_noClassKind.html"
	Case 62
        programName = "�����������"
		programTitleImg = "07"
		importFile      = "write_noClassKind.html"

	' 2014-10-28�� �߰�
	Case 89 '����ũ�����̳�
		programName = "����ũ�����̳�"
		programTitleImg = "08"
	Case 90 'Ƽ�ҹɸ���
		programName = "Ƽ�ҹɸ���"
		programTitleImg = "09"
	Case 91 '�ڵ�帳 ������
		programName = "�ڵ�帳 ������"
		programTitleImg = "10"
		importFile = "write_noClass.html"
	Case 92 'Ȩī�丶����
		programName = "Ȩī�丶����"
		programTitleImg = "11"
		importFile = "write_noClassKind.html"
End Select 


Call Expires()
Call dbopen()
	Call getData()
	If cntSubData <= -1 Then 
		With Response
		 .Write "<script language='javascript' type='text/javascript'>"
		 .Write "alert('��ϵ� ���������� �����ϴ�.');"
		 .Write "location.href='index.asp?applicationKey="&applicationKey&"';"
		 .Write "</script>"
		 .End
		End With
	End If
	Dim ProgramOption : ProgramOption = makeOption(arrSubData, cntSubData, SUB_Idx, SUB_OnData, "" )
Call dbclose()




dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "application/" & importFile ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("applicationKey", applicationKey ) _
	,array("tabOnOff1", tabOnOff1 ) _
	,array("tabOnOff2", tabOnOff2 ) _
	,array("tabOnOff3", tabOnOff3 ) _
	,array("programName", programName ) _
	,array("programTitleImg", programTitleImg ) _

	,array("ProgramKind", ProgramKind ) _
	,array("ProgramClass", ProgramClass ) _

	,array("ProgramOption", ProgramOption ) _
	,array("downlPhotos", DOWNLOAD_USER_PHOTO_PATH & FI_Photo ) _
	,array("actType", actType ) _

	,array("UserIdx", Session("UserIdx") ) _
	,array("UserName", FI_UserName ) _
	,array("UserId", FI_UserId ) _
	,array("UserBirth", FI_UserBirth ) _
	,array("UserPhone", FI_UserHphone1 &"-"&FI_UserHphone2&"-"&FI_UserHphone3 ) _
	,array("UserEmail", FI_UserEmail ) _
	,array("UserAddr", FI_UserAddr1 & "  " & UserAddr2 ) _
	,array("LastName", FI_LastName ) _
	,array("FirstName", FI_FirstName ) _
	,array("Photo", FI_Photo ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = Nothing



Sub getData()

	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @UserIdx INT;" &_
	"SET @UserIdx = ?; " &_

	"SELECT " &_
	"	 [UserName]" &_
	"	,[UserId]" &_
	"	,[UserBirth]" &_
	"	,[UserHphone1]" &_
	"	,[UserHphone2]" &_
	"	,[UserHphone3]" &_
	"	,[UserEmail]" &_
	"	,[UserAddr1]" &_
	"	,[UserAddr2]" &_
	"	,[Photo]" &_
	"	,[LastName]" &_
	"	,[FirstName]" &_
	"FROM [dbo].[SP_USER_MEMBER] WHERE [UserIdx] = @UserIdx " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,convert(varchar,[OnData],23) as [OnData] " &_
	"	,[Pay] " &_
	"FROM [dbo].[SP_PROGRAM] A " &_
	"WHERE [CodeIdx] = ?  " &_
	"AND [Dellfg] = 0  " &_
	"AND CONVERT(varchar(10),[StartDate],23) <= CONVERT(varchar(10),getDate(),23) " &_
	"AND CONVERT(varchar(10),[EndDate],23) >= CONVERT(varchar(10),getDate(),23) " &_
	"AND ISNULL([MaxNumber],0) > ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM_APP] WHERE [State] != 2 AND [ProgramIdx] = A.[Idx] ) " &_
	"ORDER BY [OnData] ASC "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		.Parameters.Append .CreateParameter( "@CodeIdx" ,adInteger , adParamInput , 0 , applicationKey )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")

	'���������� ���α׷� ���
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "SUB")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrSubData = objRs.GetRows()
		cntSubData = UBound(arrSubData, 2)
	End If

	Set objRs = Nothing
End Sub
%>