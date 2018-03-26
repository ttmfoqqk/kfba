<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrList
Dim cntList  : cntList   = -1
Dim cntTotal : cntTotal  = 0
Dim rows     : rows      = 20

Dim SHarrList
Dim SHcntList  : SHcntList   = -1

Dim pageNo     : pageNo     = RequestSet("pageNo"   , "GET" , 1)
Dim sIndate    : sIndate    = RequestSet("sIndate"  , "GET" , "")
Dim sOutdate   : sOutdate   = RequestSet("sOutdate" , "GET" , "")
Dim sOnDate    : sOnDate    = RequestSet("sOnDate"  , "GET" , "")
Dim sPcode     : sPcode     = RequestSet("sPcode"   , "GET" , "56")
Dim sArea      : sArea      = RequestSet("sArea"    , "GET" , "")

Dim sId        : sId        = RequestSet("sId"      , "GET" , "")
Dim sName      : sName      = RequestSet("sName"    , "GET" , "")
Dim sPhone3    : sPhone3    = RequestSet("sPhone3"  , "GET" , "")
Dim sState     : sState     = RequestSet("sState"   , "GET" , "")
Dim sSnumber   : sSnumber   = RequestSet("sSnumber" , "GET" , "")
Dim sKind      : sKind      = RequestSet("sKind"    , "GET" , "")
Dim sClass     : sClass     = RequestSet("sClass"   , "GET" , "")

Dim sOnTime    : sOnTime    = RequestSet("sOnTime"  , "GET" , "")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sArea="      & sArea &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sState="     & sState &_
		"&sSnumber="   & sSnumber &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass &_
		"&sOnTime="    & sOnTime

Dim pageURL
pageURL	= g_url & "?pageNo=__PAGE__" &_
		"&sIndate="    & sIndate &_
		"&sOutdate="   & sOutdate &_
		"&sOnDate="    & sOnDate &_
		"&sPcode="     & sPcode &_
		"&sArea="      & sArea &_
		"&sId="        & sId &_
		"&sName="      & sName &_
		"&sPhone3="    & sPhone3 &_
		"&sState="     & sState &_
		"&sSnumber="   & sSnumber &_
		"&sKind="      & sKind &_
		"&sClass="     & sClass &_
		"&sOnTime="    & sOnTime


checkAdminLogin(g_host & g_url & "?" & PageParams)

Dim hoursOption : hoursOption = "<option value="""">선택</option>"
for iLoop = 7 to 20
	tmp_value = IIF( iLoop < 10 , "0" & iLoop , iLoop )
	tmp_tt    = IIF( iLoop < 12 , "오전", "오후" )
	tmp_hh    = IIF( iLoop < 13 , iLoop , iLoop - 12 )
	'tmp_hh    = iLoop
	hoursOption = hoursOption & "<option value=""" & tmp_value & """" & IIF(sOnTime=Trim(tmp_value)," selected","") & ">" & tmp_tt & " " & tmp_hh & "</option>"
Next 


Call Expires()
Call dbopen()
	Call common_code_list(17) ' 프로그램명 콤보박스 옵션
	Dim codeOption : codeOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, sPcode )

	Dim StateOption
	StateOption = "" &_
	"<option value="""">선 택</option>" &_
	"<option value=""1"" "&IIF(sState="1","selected","")&" >입금대기</option>" &_
	"<option value=""0"" "&IIF(sState="0","selected","")&" >접수완료</option>" &_
	"<option value=""2"" "&IIF(sState="2","selected","")&" >접수취소</option>" &_
	"<option value=""3"" "&IIF(sState="3","selected","")&" >불합격</option>" &_
	"<option value=""4"" "&IIF(sState="4","selected","")&" >미응시(불합격)</option>" &_
	"<option value=""10"" "&IIF(sState="10","selected","")&" >합격</option>"

	Dim KindOption
	KindOption = "" &_
	"<option value="""">선 택</option>" &_
	"<option value=""1"" "&IIF(sKind="1","selected","")&" >필기</option>" &_
	"<option value=""2"" "&IIF(sKind="2","selected","")&" >실기</option>"

	Dim ClassOption
	ClassOption = "" &_
	"<option value="""">선 택</option>" &_
	"<option value=""1"" "&IIF(sClass="1","selected","")&" >1급</option>" &_
	"<option value=""2"" "&IIF(sClass="2","selected","")&" >2급</option>" &_
	"<option value=""3"" "&IIF(sClass="3","selected","")&" >3급</option>"


	Call GetList()

	Dim SearchOnDateOption : SearchOnDateOption = "<option value="""">선 택</option>"

	for iLoop = 0 to SHcntList
		SearchOnDateOption = SearchOnDateOption & "<option value=""" & SHarrList( SEARCH_OnData,iLoop) & """" & IIF(SHarrList( SEARCH_OnData,iLoop)=sOnDate," selected","") & ">" & SHarrList( SEARCH_OnData,iLoop) & "</option>"
	Next

Call dbclose()



Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageURL)

' 속도저하 문제로 프로시저로 작성
Sub GetList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("adodb.command")
	with objCmd
		.ActiveConnection	= objConn
		.prepared			= true
		.CommandType		= adCmdStoredProc
		.CommandText		= "APPLICATION_L"
		
		.Parameters("@pageNo").value   = pageNo
		.Parameters("@rows").value     = rows
		.Parameters("@sIndate").value  = sIndate
		.Parameters("@sOutdate").value = sOutdate
		.Parameters("@sOnDate").value  = sOnDate
		.Parameters("@sPcode").value   = sPcode
		.Parameters("@sArea").value    = sArea
		.Parameters("@sId").value      = sId
		.Parameters("@sName").value    = sName
		.Parameters("@sPhone3").value  = sPhone3
		.Parameters("@sState").value   = sState
		.Parameters("@sSnumber").value = sSnumber
		.Parameters("@sKind").value    = sKind
		.Parameters("@sClass").value   = sClass
		.Parameters("@sOnTime").value  = sOnTime
		set objRs = .Execute
	End with
	set objCmd = Nothing
	
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
		cntTotal	= arrList(FI_tcount, 0)
	End If

	'검정일자 검색용 셀렉트
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "SEARCH")
	If Not(objRs.Eof or objRs.Bof) Then		
		SHarrList = objRs.GetRows()
		SHcntList = UBound(SHarrList, 2)
	End If

	objRs.close	: Set objRs = Nothing
End Sub

dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "application/leftMenu.html"
ntpl.setFile "MAIN", "application/applicationL.html"
ntpl.setFile "FOOTER", "_inc/footer.html"

'//상단메뉴오버
Call topMenuOver()


call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST1","LEFT_MENU_LIST2"))
If common_code_cntList > -1 Then 
	for iLoop = 0 to common_code_cntList
		ntpl.setBlockReplace array( _
			 array("Idx" , common_code_arrList(CCODE_Idx,iLoop)  ) _
			,array("Name", common_code_arrList(CCODE_Name,iLoop) ) _
			,array("leftMenuOverClass", IIF( CStr(common_code_arrList(CCODE_Idx,iLoop))=sPcode,"admin_left_over","" ) ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
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

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("LEFT_MENU_LIST2")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST2")
End If

call ntpl.setBlock("MAIN", array("APPLICATION_LOOP","LOOP_NODATA"))
'// BLOCK 부분 처리

If cntList > -1 Then 
	for iLoop = 0 to cntList

	StateTxt = ""

	If arrList(FI_State,iLoop) = "0" Then 
		StateTxt = "접수완료"
	ElseIf arrList(FI_State,iLoop) = "1" Then 
		StateTxt = "<font color=""#11179a"">입금대기</font>"
	ElseIf arrList(FI_State,iLoop) = "2" Then 
		StateTxt = "<font color=""#9a1134"">접수취소</font>"
	ElseIf arrList(FI_State,iLoop) = "3" Then 
		StateTxt = "<font color=""#9a1134"">불합격</font>"
	ElseIf arrList(FI_State,iLoop) = "4" Then 
		StateTxt = "<font color=""#9a1134"">미응시(불합격)</font>"
	ElseIf arrList(FI_State,iLoop) = "10" Then 
		StateTxt = "<font color=""#11179a"">합격</font>"
	End If

	PrograName = arrList(FI_ProgramNema,iLoop)

	If arrList(FI_Class,iLoop) = "1" Then
		PrograName = PrograName & " 1급"
	ElseIf arrList(FI_Class,iLoop) = "2" Then
		PrograName = PrograName & " 2급"
	ElseIf arrList(FI_Class,iLoop) = "3" Then
		PrograName = PrograName & " 3급"
	End If

	If arrList(FI_Kind,iLoop) = "1" Then
		PrograName = PrograName & " [필기]"
	ElseIf arrList(FI_Kind,iLoop) = "2" Then
		PrograName = PrograName & " [실기]"
	End If

	StateMyOption = "" &_
	"<option value=""1"" "&IIF(arrList(FI_State,iLoop)="1","selected","")&" >입금대기</option>" &_
	"<option value=""0"" "&IIF(arrList(FI_State,iLoop)="0","selected","")&" >접수완료</option>" &_
	"<option value=""2"" "&IIF(arrList(FI_State,iLoop)="2","selected","")&" >접수취소</option>" &_
	"<option value=""3"" "&IIF(arrList(FI_State,iLoop)="3","selected","")&" >불합격</option>" &_
	"<option value=""4"" "&IIF(arrList(FI_State,iLoop)="4","selected","")&" >미응시</option>" &_
	"<option value=""10"" "&IIF(arrList(FI_State,iLoop)="10","selected","")&" >합격</option>"


		ntpl.setBlockReplace array( _
			 array("rownum" , arrList(FI_rownum,iLoop)  ) _
			,array("Idx" , arrList(FI_Idx,iLoop)  ) _
			,array("UserId", arrList(FI_UserId,iLoop) ) _
			,array("UserName", arrList(FI_UserName,iLoop) ) _
			,array("UserPhone", arrList(FI_UserHphone1,iLoop) &"-"& arrList(FI_UserHphone2,iLoop) &"-"& arrList(FI_UserHphone3,iLoop) ) _
			,array("ProgramNema" , PrograName  ) _
			,array("AreaName" , arrList(FI_AreaName,iLoop)  ) _
			,array("OnData" , arrList(FI_OnData,iLoop)  ) _
			,array("InData" , arrList(FI_InData,iLoop)  ) _
			,array("StateMyOption" , StateMyOption  ) _
			,array("State" , arrList(FI_State,iLoop)  ) _

			,array("Snumber" , IIF(arrList(FI_Snumber,iLoop)="","&nbsp;",arrList(FI_Snumber,iLoop))  ) _
		), ""

		'// MEMBER_LOOP 블럭 누적
		ntpl.tplParseBlock("APPLICATION_LOOP")
	Next
	ntpl.tplBlockDel("LOOP_NODATA")
Else
	ntpl.tplBlockDel("APPLICATION_LOOP")
	ntpl.tplParseBlock("LOOP_NODATA")
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("pageList", pagelist ) _
	,array("PageParams", PageParams ) _

	,array("codeOption", codeOption ) _
	,array("StateOption", StateOption ) _
	

	,array("pageNo"    , pageNo ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sOnDate" , sOnDate ) _
	,array("sPcode", sPcode ) _
	,array("sArea", sArea ) _
	,array("sId", sId ) _
	,array("sName", sName ) _
	,array("sPhone3", sPhone3 ) _
	,array("sState", sState ) _
	,array("sOnTime", sOnTime ) _
	
	,array("SearchOnDateOption", SearchOnDateOption ) _
	,array("sSnumber", sSnumber ) _
	,array("KindOption", KindOption ) _
	,array("ClassOption", ClassOption ) _
	,array("hoursOption", hoursOption ) _
	

	,array("s1Day"    , Date() ) _
	,array("s7Day"    , Date() -7 ) _
	,array("s30Day"   , Date() -30 ) _

	,array("leftMenuOverClass1"   , "admin_left_over" ) _
	,array("leftMenuOverClass2"   , "" ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = nothing

%>