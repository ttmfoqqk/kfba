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
Dim applicationKey  : applicationKey  = RequestSet("applicationKey" ,"GET",56)
Dim LGD_FINANCENAME : LGD_FINANCENAME = RequestSet("LGD_FINANCENAME","GET","")
Dim LGD_ACCOUNTNUM  : LGD_ACCOUNTNUM  = RequestSet("LGD_ACCOUNTNUM" ,"GET","")
Dim tabOnOff1 : tabOnOff1 = "_off"
Dim tabOnOff2 : tabOnOff2 = "_off"
Dim tabOnOff3 : tabOnOff3 = "_on"

Dim actType   : actType   = "INSERT"

Dim programName

Select Case applicationKey
    Case 56
        programName = "Ŀ�ǹٸ���Ÿ"
		programTitleImg = "01"
    Case 57
        programName = "Ĭ�������ֻ�"
		programTitleImg = "02"
    Case 58
        programName = "�ͼַ�����Ʈ"
		programTitleImg = "03"
	Case 59
        programName = "���μҹɸ���"
		programTitleImg = "04"
	Case 60
        programName = "�����ְ�����"
		programTitleImg = "05"
	Case 61
        programName = "�ܽİ濵������"
		programTitleImg = "06"
	Case 62
        programName = "�����������"
		programTitleImg = "07"

	' 2014-10-28�� �߰�
	Case 89 '����ũ�����̳�
		programName = "����ũ�����̳�"
		programTitleImg = "08"
	Case 90 'Ƽ�ҹɸ���
		programName = "Ƽ�ҹɸ���"
		programTitleImg = "09"
	Case 91 '���Ḷ����
		programName = "���Ḷ����"
		programTitleImg = "10"
	Case 92 'Ȩī�丶����
		programName = "Ȩī�丶����"
		programTitleImg = "11"
End Select 


dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "application/payBank.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// ��� �α��� ��ó��
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")


ntpl.tplAssign array(   _
	 array("imgDir"         , TPL_DIR_IMAGES ) _
	,array("applicationKey" , applicationKey ) _
	,array("tabOnOff1"      , tabOnOff1 ) _
	,array("tabOnOff2"      , tabOnOff2 ) _
	,array("tabOnOff3"      , tabOnOff3 ) _
	,array("programName"    , programName ) _
	,array("LGD_FINANCENAME", LGD_FINANCENAME ) _
	,array("LGD_ACCOUNTNUM" , LGD_ACCOUNTNUM ) _
	,array("programTitleImg", programTitleImg ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = Nothing
%>