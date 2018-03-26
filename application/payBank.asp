<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
'56 : 커피바리스타
'57 : 칵테일조주사
'58 : 믹솔로지스트
'59 : 와인소믈리에
'60 : 전통주관리사
'61 : 외식경영관리사
'62 : 식음료관리사

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
        programName = "커피바리스타"
		programTitleImg = "01"
    Case 57
        programName = "칵테일조주사"
		programTitleImg = "02"
    Case 58
        programName = "믹솔로지스트"
		programTitleImg = "03"
	Case 59
        programName = "와인소믈리에"
		programTitleImg = "04"
	Case 60
        programName = "전통주관리사"
		programTitleImg = "05"
	Case 61
        programName = "외식경영관리사"
		programTitleImg = "06"
	Case 62
        programName = "식음료관리사"
		programTitleImg = "07"

	' 2014-10-28일 추가
	Case 89 '케이크디자이너
		programName = "케이크디자이너"
		programTitleImg = "08"
	Case 90 '티소믈리에
		programName = "티소믈리에"
		programTitleImg = "09"
	Case 91 '음료마스터
		programName = "음료마스터"
		programTitleImg = "10"
	Case 92 '홈카페마스터
		programName = "홈카페마스터"
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
'// 상단 로그인 블럭처리
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

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing
%>