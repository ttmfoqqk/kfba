<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim Idx      : Idx      = RequestSet("Idx"    ,"GET",0)
Dim pageNo   : pageNo   = RequestSet("pageNo" ,"GET",1 )
Dim sPcode   : sPcode   = RequestSet("sPcode" ,"GET","")
Dim sACode   : sACode   = RequestSet("sACode" ,"GET","")

Dim PageParams
PageParams = "pageNo=" & pageNo &_
		"&sPcode=" & sPcode &_
		"&sACode=" & sACode

Call Expires()
Call dbopen()
	Call getData()
Call dbclose()

PhotoExt = FILE_CHECK_EXT_RETURN( FI_Map )
If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
	MapImages = img_resize( "/upload/programsArea/", FI_Map ,600,600)
Else
	MapImages= "<a href=""_lib/dowload.asp?pach=/upload/programsArea/&file="&FI_Map&""">"& FI_Map &"</a>"
End If

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "map/view.html" ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")


ntpl.tplAssign array(   _
	 array("imgDir" , TPL_DIR_IMAGES ) _
	,array("sPcode" , sPcode ) _
	,array("sACode" , sACode ) _
	,array("pageNo" , pageNo ) _
	,array("PageParams", PageParams ) _
	,array("Idx", FI_Idx ) _
	,array("Name", TagDecode(Trim( FI_Name )) ) _
	,array("Addr", TagDecode(Trim( FI_Addr )) ) _
	,array("Tel", TagDecode(Trim( FI_Tel )) ) _
	,array("Info", Trim( FI_Info ) ) _
	,array("WebAddr", TagDecode(Trim( FI_WebAddr )) ) _
	,array("AreaWebUrl"  , IIF( FI_WebAddr="","&nbsp;", "<a href="""&Replace(FI_WebAddr,"http://","")&""" target=""_blank"">"&FI_WebAddr&"</a> " ) ) _
	,array("Map", FI_Map ) _
	,array("Code", FI_Code ) _
	,array("MapImages", MapImages ) _
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
	"DECLARE @Idx INT;" &_
	"SET @Idx = ?; " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,[Name] " &_
	"	,[Addr] " &_
	"	,[Tel] " &_
	"	,[Info] " &_
	"	,[WebAddr] " &_
	"	,[Map] " &_
	"	,[Code] " &_
	"	,[CodeIdx] " &_
	"	,[AddrIdx] " &_
	"FROM [dbo].[SP_PROGRAM_AREA] " &_
	"WHERE [Idx] = @Idx AND [Dellfg] = 0"

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , Idx )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub
%>