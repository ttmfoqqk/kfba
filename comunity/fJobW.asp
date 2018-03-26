<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%


Dim arrCareer  , arrQualify
Dim cntCareer  : cntCareer  = -1
Dim cntQualify : cntQualify = -1

Dim Idx      : Idx      = RequestSet("Idx"      ,"GET",0)
Dim pageNo   : pageNo   = RequestSet("pageNo"   ,"GET",1)
Dim sName    : sName    = RequestSet("sName"    ,"GET",0)
Dim sId      : sId      = RequestSet("sId"      ,"GET",0)
Dim sTitle   : sTitle   = RequestSet("sTitle"   ,"GET",0)
Dim sContant : sContant = RequestSet("sContant" ,"GET",0)
Dim sWord    : sWord    = RequestSet("sWord"    ,"GET","")
Dim actType  : actType  = RequestSet("actType"  ,"GET","")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&BoardKey=" & BoardKey &_
		"&sName="    & sName &_
		"&sId="      & sId &_
		"&sTitle="   & sTitle &_
		"&sContant=" & sContant &_
		"&sWord="    & sWord

checkLogin( g_host & g_url &"?"&PageParams )



Call Expires()
Call dbopen()
	Call getData()
	actType = IIF( FI_Idx = "","INSERT", IIF( actType="","MODIFY",actType ) )

	Call common_code_list(18) ' 형태
	Dim FormOption : FormOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, int(IIF(FI_Form="",0,FI_Form)) )	
	Call common_code_list(19) ' 직종
	Dim KindOption   : KindOption = makeOption(common_code_arrList, common_code_cntList, CCODE_Idx, CCODE_Name, int(IIF(FI_Kind="",0,FI_Kind)) )

Call dbclose()



dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "comunity/fJobW.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("CAREER_LOOP" , "CAREER_LOOP_NODATA" , "QUALIFY_LOOP" , "QUALIFY_LOOP_NODATA"))
If cntCareer > -1 Then 

	for iLoop = 0 to cntCareer
		ntpl.setBlockReplace array( _
			 array("Name", Trim( arrCareer(CAREER_Name,iLoop) ) )_
			,array("WorkMonth", Trim( arrCareer(CAREER_WorkMonth,iLoop) ) )_
			,array("LastPosition", Trim( arrCareer(CAREER_LastPosition,iLoop) ) )_
			
		), ""
		ntpl.tplParseBlock("CAREER_LOOP")
	Next 
	ntpl.tplBlockDel("CAREER_LOOP_NODATA")
Else
	ntpl.tplParseBlock("CAREER_LOOP_NODATA")
	ntpl.tplBlockDel("CAREER_LOOP")
End If

If cntQualify > -1 Then 

	for iLoop = 0 to cntQualify
		ntpl.setBlockReplace array( _
			 array("Name", Trim( arrQualify(Qualify_Name,iLoop) ) )_
			,array("Tdate", trim( arrQualify(Qualify_Tdate,iLoop) ) )_
			,array("Publish", Trim( arrQualify(Qualify_Publish,iLoop) ) )_
			
		), ""
		ntpl.tplParseBlock("QUALIFY_LOOP")
	Next 
	ntpl.tplBlockDel("QUALIFY_LOOP_NODATA")
Else
	ntpl.tplParseBlock("QUALIFY_LOOP_NODATA")
	ntpl.tplBlockDel("QUALIFY_LOOP")
End If


ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("BoardName" , BDV_Name ) _
	,array("actType"   , actType ) _
	,array("pageNo"    , pageNo ) _
	,array("sName"     , sName ) _
	,array("sId"       , sId ) _
	,array("sTitle"    , sTitle ) _
	,array("sContant"  , sContant ) _
	,array("sWord"     , sWord ) _
	,array("PageParams", PageParams ) _


	,array("UserName", USER_UserName ) _
	,array("UserBirth", USER_UserBirth ) _
	,array("UserPhone", USER_UserHphone1&"-"&USER_UserHphone2&"-"&USER_UserHphone3 ) _
	,array("UserEmail", USER_UserEmail ) _
	,array("Photo", USER_Photo ) _

	,array("Idx", FI_Idx ) _
	,array("FormOption", FormOption ) _
	,array("KindOption", KindOption ) _
	,array("WorkArea", FI_WorkArea ) _
	,array("Pay", FI_Pay ) _
	,array("School", FI_School ) _
	,array("Bigo", FI_Bigo ) _
	,array("File", FI_File ) _
	,array("InData", FI_Indate ) _
	,array("Ip", FI_Ip ) _
	,array("downloadUrl", DOWNLOAD_BASE_PATH_JOB & FI_File ) _
	,array("downlPhotos", DOWNLOAD_USER_PHOTO_PATH & USER_Photo ) _

	,array("UserIdx", FI_UserIdx ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing





Sub getData()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Idx INT,@UserIdx INT;" &_
	"SET @Idx = ?; " &_
	"SET @UserIdx = ?; " &_
	"SELECT " &_
	"	 [Idx] " &_
	"	,[Form] " &_
	"	,[Kind] " &_
	"	,[WorkArea] " &_
	"	,[Pay] " &_
	"	,[School] " &_
	"	,[Bigo] " &_
	"	,[File] " &_
	"	,[InData] " &_
	"	,[Ip] " &_
	"	,[UserIdx]" &_
	"FROM [dbo].[SP_JOB_USER] " &_
	"WHERE [Idx] = @Idx AND [UserIdx] = @UserIdx AND [Dellfg] = 0 " &_

	"SELECT " &_
	"	 [UserName]" &_
	"	,[UserBirth]" &_
	"	,[UserHphone1]" &_
	"	,[UserHphone2]" &_
	"	,[UserHphone3]" &_
	"	,[UserEmail]" &_
	"	,[Photo]" &_
	"FROM [dbo].[SP_USER_MEMBER] WHERE [UserIdx] = @UserIdx " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,[Name] " &_
	"	,[WorkMonth] " &_
	"	,[LastPosition] " &_
	"	,[UserIdx] " &_
	"FROM [dbo].[SP_JOB_USER_CAREER] WHERE [UserIdx] = @UserIdx " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,[Name] " &_
	"	,[Tdate] " &_
	"	,[Publish] " &_
	"	,[UserIdx] " &_
	"FROM [dbo].[SP_JOB_USER_QUALIFY] WHERE [UserIdx] = @UserIdx "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , Idx )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	
	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "USER")

	'경력사항
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "CAREER")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrCareer = objRs.GetRows()
		cntCareer = UBound(arrCareer, 2)
	End If

	'자격증
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "Qualify")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrQualify = objRs.GetRows()
		cntQualify = UBound(arrQualify, 2)
	End If

	Set objRs = Nothing
End Sub
%>