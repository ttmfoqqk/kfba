<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList : cntList  = -1

Dim Idx        : Idx        = RequestSet("Idx" , "GET" , 0)
Dim pageNo     : pageNo     = RequestSet("pageNo" , "GET" , 1)

Dim sOnDate    : sOnDate    = RequestSet("sOnDate"  , "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "56")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sKind      : sKind      = RequestSet("sKind"    , "GET" , "")
Dim sClass     : sClass     = RequestSet("sClass"   , "GET" , "")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sName="      & sName &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass

checkAdminLogin(g_host & g_url & "?" & PageParams & "&Idx=" & Idx)

Dim importFile  : importFile  = "programsW.html"

Call Expires()
Call dbopen()
	Call getView()
	Call common_code_list(17) ' ���α׷��� �޺��ڽ� �ɼ�
	Dim codeOption  : codeOption  = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, int(IIF(FV_CodeIdx="", IIF( sPcode="",0,sPcode ) ,FV_CodeIdx)) )
	Dim KindOption  : KindOption  = "<option value="""">�������</option><option value=""1"""&IIF(FV_Kind="1"," selected","")&">�ʱ�</option><option value=""2"""&IIF(FV_Kind="2"," selected","")&">�Ǳ�</option>"
	Dim ClassOption : ClassOption = "<option value="""">�޼�</option><option value=""1"""&IIF(FV_Class="1"," selected","")&">1��</option><option value=""2"""&IIF(FV_Class="2"," selected","")&">2��</option>"

	Dim hoursOption : hoursOption = "<option value="""">����</option>"
	for iLoop = 7 to 20
		tmp_value = IIF( iLoop < 10 , "0" & iLoop , iLoop )
		tmp_tt    = IIF( iLoop < 12 , "����", "����" )
		tmp_hh    = IIF( iLoop < 13 , iLoop , iLoop - 12 )
		'tmp_hh    = iLoop
		hoursOption = hoursOption & "<option value=""" & tmp_value & """" & IIF(FV_OnDataHours=Trim(tmp_value)," selected","") & ">" & tmp_tt & " " & tmp_hh & "</option>"
	Next 

	Dim minutesOption : minutesOption = "<option value="""">����</option>"
	for iLoop = 0 to 59 Step 10
		tmp_value = IIF( iLoop < 10 , "0" & iLoop , iLoop )
		minutesOption = minutesOption & "<option value=""" & tmp_value & """" & IIF(FV_OnDataMinutes=Trim(tmp_value)," selected","") & ">" & iLoop & "</option>"
	Next 

	Select Case IIF(sPcode="",0,sPcode)
		Case 56 'Ŀ�ǹٸ���Ÿ
			ClassOption = ClassOption & "<option value=""3"""&IIF(FV_Class="3"," selected","")&">3��</option>"
		Case 57 'Ĭ�������ֻ�
			importFile = "programsW_noClass.html"
		Case 58 '�ͼַ�����Ʈ

		Case 59 '���μҹɸ���

		Case 60 '�󶼾�Ʈ ������
			importFile = "programsW_noClassKind.html"
		Case 61 '�ܽİ濵������
			importFile = "programsW_noClassKind.html"
		Case 62 '�����������
			importFile = "programsW_noClassKind.html"

		' 2014-10-28�� �߰�
		Case 89 '����ũ�����̳�

		Case 90 'Ƽ�ҹɸ���

		Case 91 '�ڵ�帳 ������
			importFile = "programsW_noClass.html"
		Case 92 'Ȩī�丶����
			importFile = "programsW_noClassKind.html"
	End Select 

Call dbclose()

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @Idx INT; SET @Idx = ? ;" &_
	"DECLARE @Pcode INT; SET @Pcode = ? ;" &_
	"SELECT " &_
	"	 [Idx] " &_
	"	,[CodeIdx] " &_
	"	,convert(varchar,[OnData],23) AS [OnData] " &_
	"	,SUBSTRING(CONVERT(VARCHAR(8), [OnData], 108),1,2) AS [OnDataHours] " &_
	"	,SUBSTRING(CONVERT(VARCHAR(8), [OnData], 108),4,2) AS [OnDataMinutes] " &_
	"	,convert(varchar,[StartDate],23) AS [StartDate] " &_
	"	,convert(varchar,[EndDate],23) AS [EndDate] " &_
	"	,[Pay] " &_
	"	,ISNULL([MaxNumber],0) AS [MaxNumber] " &_
	"	,[Kind] " &_
	"	,[Class] " &_
	"	,[InDate] " &_
	"FROM [dbo].[SP_PROGRAM] " &_
	"WHERE [Idx] = @Idx AND [Dellfg] = 0 " &_

	"SELECT " &_
	"	 A.[Idx] " &_
	"	,A.[Name] " &_
	"	,A.[Addr] " &_
	"	,A.[Tel] " &_
	"	,A.[Info] " &_
	"	,A.[WebAddr] " &_
	"	,A.[Map] " &_
	"	,A.[Dellfg] " &_
	"	,B.[AreaIdx] AS [check] " &_
	"FROM [dbo].[SP_PROGRAM_AREA] A " &_
	"LEFT JOIN [dbo].[SP_PROGRAM_ON_AREA] B ON(A.[Idx] = B.[AreaIdx] AND B.[ProgramIdx] = @Idx) " &_
	"WHERE A.[Dellfg] = 0 " &_
	"AND ( [Code] IS NOT NULL OR [Code] !='' ) " &_
	"AND [CodeIdx] = @Pcode " &_
	"ORDER BY [Idx] DESC "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput, 0, Idx )
		.Parameters.Append .CreateParameter( "@Pcode" ,adInteger , adParamInput, 0, sPcode )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")


	'������ ���
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "FULL")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrList = objRs.GetRows()
		cntList = UBound(arrList, 2)
	End If

	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// ���ø� ���丮 ���� (�⺻ tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "programs/leftMenu.html"
ntpl.setFile "MAIN", "programs/" & importFile
ntpl.setFile "FOOTER", "_inc/footer.html"

'//��ܸ޴�����
Call topMenuOver()

call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST1","LEFT_MENU_LIST2"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("sKey", common_code_arrList(CCODE_sKey,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(common_code_arrList(CCODE_Idx,iLoop))=sPcode,"admin_left_over","" ) ) _
		), ""

		'// MEMBER_LOOP �� ����
		ntpl.tplParseBlock("LEFT_MENU_LIST1")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST1")
End If

If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
		), ""

		'// MEMBER_LOOP �� ����
		ntpl.tplParseBlock("LEFT_MENU_LIST2")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST2")
End If

call ntpl.setBlock("MAIN", array("AREA_LOOP"))
If cntList > -1 Then

	for iLoop = 0 to cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , arrList(FULL_Idx,iLoop) )  _
			,array("Name", arrList(FULL_Name,iLoop) ) _
			,array("Addr", arrList(FULL_Addr,iLoop) ) _
			,array("Tel" , arrList(FULL_Tel,iLoop) ) _
			,array("checked", IIF( Isnull(arrList(FULL_check,iLoop)) Or arrList(FULL_check,iLoop)="","","checked" ) ) _
		), ""
		ntpl.tplParseBlock("AREA_LOOP")
	Next 

Else
	ntpl.tplBlockDel("AREA_LOOP")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("PageParams", PageParams ) _

	,array("pageNo"  , pageNo ) _
	,array("sOnDate" , sOnDate ) _
	,array("sPcode"  , sPcode ) _
	,array("sName"   , sName ) _
	,array("sKind"   , sKind ) _
	,array("sClass"  , sClass ) _

	,array("actType", IIF( FV_Idx="","INSERT","UPDATE") ) _
	,array("Idx", FV_Idx ) _
	,array("codeOption", codeOption ) _
	,array("KindOption", KindOption ) _
	,array("ClassOption", ClassOption ) _

	,array("hoursOption", hoursOption ) _
	,array("minutesOption", minutesOption ) _

	,array("OnData", FV_OnData ) _
	,array("Pay", IIF( FV_Pay="",0,FV_Pay ) ) _
	,array("InDate", FV_InDate ) _

	,array("StartDate", FV_StartDate ) _
	,array("EndDate", FV_EndDate ) _
	,array("MaxNumber", FV_MaxNumber ) _

	,array("leftMenuOverClass1"   , "admin_left_over" ) _
	,array("leftMenuOverClass2"   , "" ) _
), ""

'// �������� { ��ũ ����� ���� ��
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// ������ ���ø� ����ó��
ntpl.tplPrint()  '// ���

set ntpl = nothing

%>