<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../../_lib/template.class.asp" -->
<!-- #include file = "../../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%
Dim arrCareer  , arrQualify
Dim cntCareer  : cntCareer  = -1
Dim cntQualify : cntQualify = -1

Dim Idx          : Idx          = RequestSet("Idx"        ,"GET",0)
Dim pageNo       : pageNo       = RequestSet("pageNo","GET",1)
Dim actType      : actType      = RequestSet("actType","GET","")


Dim sIndate      : sIndate      = RequestSet("sIndate"     ,"GET","")
Dim sOutdate     : sOutdate     = RequestSet("sOutdate"    ,"GET","")
Dim sId          : sId          = RequestSet("sId"         ,"GET","")
Dim sName        : sName        = RequestSet("sName"       ,"GET","")


Call Expires()
Call dbopen()
	Call BoardCodeList()
	If IsNumeric(Idx) Then 
		Call getData()
	End If
Call dbclose()

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sIndate="  & sIndate &_
		"&sOutdate=" & sOutdate &_
		"&sId="      & sId &_
		"&sName="    & sName

checkAdminLogin(g_host & g_url & "?" & PageParams  & "&Idx=" & Idx)


dim ntpl
set ntpl = new SkyTemplate

'// 템플릿 디렉토리 설정 (기본 tpl)
ntpl.setTplDir( ADMIN_ROOT_DIR & TPL_DIR_FOLDER )
ntpl.setFile "HEADER", "_inc/header.html"
ntpl.setFile "LEFT", "customer/leftMenu.html"
ntpl.setFile "MAIN", "customer/fJObV.html"
ntpl.setFile "FOOTER", "_inc/footer.html"
'//상단메뉴오버
Call topMenuOver()

'//왼쪽메뉴 설정
call ntpl.setBlock("LEFT", array("LEFT_MENU_LIST"))
If BC_CNT_LIST > -1 Then 
	for iLoop = 0 to BC_CNT_LIST
		ntpl.setBlockReplace array( _
			array("Idx", BC_ARRY_LIST(BDL_Idx,iLoop) ), _
			array("Name", BC_ARRY_LIST(BDL_Name,iLoop) ) _
		), ""
		ntpl.tplParseBlock("LEFT_MENU_LIST")
	Next
Else
	ntpl.tplBlockDel("LEFT_MENU_LIST")
End If
'//왼쪽메뉴 설정 끝

call ntpl.setBlock("MAIN", array("CAREER_LOOP" , "QUALIFY_LOOP" ))

If cntCareer > -1 Then 

	for iLoop = 0 to cntCareer
		ntpl.setBlockReplace array( _
			 array("Name", arrCareer(CAREER_Name,iLoop) )_
			,array("WorkMonth", arrCareer(CAREER_WorkMonth,iLoop) )_
			,array("LastPosition", arrCareer(CAREER_LastPosition,iLoop) )_
			
		), ""
		ntpl.tplParseBlock("CAREER_LOOP")
	Next 
Else
	ntpl.tplBlockDel("CAREER_LOOP")
End If

If cntQualify > -1 Then 

	for iLoop = 0 to cntQualify
		ntpl.setBlockReplace array( _
			 array("Name", arrQualify(Qualify_Name,iLoop) )_
			,array("Tdate", arrQualify(Qualify_Tdate,iLoop) )_
			,array("Publish", arrQualify(Qualify_Publish,iLoop) )_
			
		), ""
		ntpl.tplParseBlock("QUALIFY_LOOP")
	Next 
Else
	ntpl.tplBlockDel("QUALIFY_LOOP")
End If

Dim PhotoExt : PhotoExt = FILE_CHECK_EXT_RETURN(FI_Photo)

If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
	UserPhotos = img_resize(USER_PHOTO_PATH,FI_Photo,150,200)
Else
	UserPhotos= "<a href="""&DOWNLOAD_USER_PHOTO_PATH & FI_Photo&""">"&FI_Photo&"</a>"
End If

ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("BoardName" , BDV_Name ) _
	,array("actType"   , actType ) _
	,array("pageNo"    , pageNo ) _
	,array("pagelist"  , pagelist ) _
	,array("sIndate"   , sIndate ) _
	,array("sOutdate"  , sOutdate ) _
	,array("sId"       , sId ) _
	,array("sName"     , sName ) _
	,array("PageParams", PageParams ) _

	,array("UserName" , FI_UserName ) _
	,array("UserBirth", FI_UserBirth ) _
	,array("UserPhone", FI_UserHphone1&"-"&FI_UserHphone2&"-"&FI_UserHphone3 ) _
	,array("UserEmail", FI_UserEmail ) _
	,array("Photo", FI_Photo ) _

	,array("Idx", FI_Idx ) _
	,array("Form", FI_Form ) _
	,array("Kind", FI_Kind ) _
	,array("WorkArea", FI_WorkArea ) _
	,array("Pay", FI_Pay ) _
	,array("School", FI_School ) _
	,array("Bigo", FI_Bigo ) _
	,array("File", FI_File ) _
	,array("InData", FI_Indate ) _
	,array("Ip", FI_Ip ) _
	,array("downloadUrl", DOWNLOAD_BASE_PATH_JOB & FI_File ) _
	,array("Photos", UserPhotos ) _

	,array("leftMenuOverClass1"   , "" ) _
	,array("leftMenuOverClass2"   , "admin_left_over" ) _
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
	"DECLARE @UserIdx INT,@Idx INT;" &_
	"SET @Idx = ?; " &_
	"SET @UserIdx = (SELECT [UserIdx] FROM [dbo].[SP_JOB_USER] WHERE [Idx] = @Idx ) ; " &_

	"SELECT " &_
	"	 A.[Idx] " &_
	"	,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Form] ) AS [Form] " &_
	"	,(SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [Kind] ) AS [Kind] " &_
	"	,A.[WorkArea] " &_
	"	,A.[Pay] " &_
	"	,A.[School] " &_
	"	,A.[Bigo] " &_
	"	,A.[File] " &_
	"	,A.[InData] " &_
	"	,A.[Ip] " &_
	"	,B.[UserName]" &_
	"	,B.[UserBirth]" &_
	"	,B.[UserHphone1]" &_
	"	,B.[UserHphone2]" &_
	"	,B.[UserHphone3]" &_
	"	,B.[UserEmail]" &_
	"	,B.[Photo]" &_
	"FROM [dbo].[SP_JOB_USER] A " &_
	"INNER JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx])" &_
	"WHERE [Idx] = @Idx AND [Dellfg] = 0 " &_

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
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")

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