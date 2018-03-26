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
Dim arrList
Dim cntList : cntList = -1

Dim arrListSub
Dim cntListSub : cntListSub = -1

Dim applicationKey : applicationKey = RequestSet("applicationKey","GET",56)
Dim tabOnOff1 : tabOnOff1 = "_off"
Dim tabOnOff2 : tabOnOff2 = "_on"
Dim tabOnOff3 : tabOnOff3 = "_off"

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
End Select 


Call Expires()
Call dbopen()
	Call getList()
Call dbclose()


dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "application/judgeL.html") _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")
call ntpl.setBlock("MAIN", array("JUDGE_LOOF","JUDGE_LOOF_NODATA"))




If cntList > -1 Then 
	Call Expires()
	Call dbopen()

	Dim PhotoExt
	for iLoop = 0 to cntList

		tmp_sub_html = ""
		Call getListSub( arrList(FI_UserIdx,iLoop) )
		If cntListSub > -1 Then
			for iLoop2 = 0 to cntListSub
			tmp_sub_html = tmp_sub_html & "<tr>" &_
			"	<td class=""cell_view_cont"">"& arrListSub(SUB_CompanyName,iLoop2) &"</td>" &_
			"	<td class=""cell_view_cont"">"& arrListSub(SUB_WorkTime,iLoop2) &"</td>" &_
			"	<td class=""cell_view_cont"">"& arrListSub(SUB_WorkMonth,iLoop2) &"</td>" &_
			"	<td class=""cell_view_cont"">"& arrListSub(SUB_LastPosition,iLoop2) &"</td>" &_
			"</tr>"
			Next
		End If

		

		PhotoExt = FILE_CHECK_EXT_RETURN( arrList(FI_Photo,iLoop) )
		If PhotoExt = "jpg" Or PhotoExt = "jpeg" Or PhotoExt = "gif" Or PhotoExt = "png" Or PhotoExt = "bmp" Then
			UserPhotos = img_resize(USER_PHOTO_PATH, arrList(FI_Photo,iLoop) ,150,200)
		Else
			UserPhotos= "<a href="""&DOWNLOAD_FI_Photo_PATH &  arrList(FI_Photo,iLoop) &""">"& arrList(FI_Photo,iLoop) &"</a>"
		End If

		
		ntpl.setBlockReplace array( _
			 array("UserName", arrList(FI_UserName,iLoop) )_
			,array("UserBirth", arrList(FI_UserBirth,iLoop) )_
			,array("UserEmail", arrList(FI_UserEmail,iLoop) )_
			,array("Photo", UserPhotos )_

			,array("JUDGE_LOOF_SUB", tmp_sub_html )_
		), ""
		ntpl.tplParseBlock("JUDGE_LOOF")
		
		
		
		
	
	Next 
	Call dbclose()
	ntpl.tplBlockDel("JUDGE_LOOF_NODATA")
Else
	ntpl.tplParseBlock("JUDGE_LOOF_NODATA")
	ntpl.tplBlockDel("JUDGE_LOOF")
End If




ntpl.tplAssign array(   _
	 array("imgDir", TPL_DIR_IMAGES ) _
	,array("applicationKey", applicationKey ) _
	,array("tabOnOff1", tabOnOff1 ) _
	,array("tabOnOff2", tabOnOff2 ) _
	,array("tabOnOff3", tabOnOff3 ) _
	,array("programName", programName ) _
	,array("programTitleImg", programTitleImg ) _
	
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing


Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; "&_
	"SELECT "&_
	"	 B.[UserIdx]" &_
	"	,B.[UserName]" &_
	"	,B.[UserBirth]" &_
	"	,B.[UserEmail]" &_
	"	,B.[Photo]" &_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP] A "&_
	"INNER JOIN [dbo].[SP_USER_MEMBER] B ON(A.[UserIdx] = B.[UserIdx])" &_
	"WHERE A.[State] = 0 AND B.[UserDelFg] = 0 AND A.[Dellfg] = 0 " &_
	"AND A.[ProgramIdx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@applicationKey" ,adInteger , adParamInput, 0, applicationKey )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "FI")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
	End If
	Set objRs = Nothing
End Sub


Sub getListSub( UserIdx )

	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; "&_
	"SELECT "&_
	"	 [CompanyName] "&_
	"	,[WorkTime] "&_
	"	,[WorkMonth] "&_
	"	,[LastPosition] "&_
	"	,[UserIdx] "&_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP_CAREER] "&_
	"WHERE [UserIdx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , UserIdx )
		set objRs = .Execute
	End With
	call cmdclose()
	CALL setFieldIndex(objRs, "SUB")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrListSub = objRs.GetRows()
		cntListSub = UBound(arrListSub, 2)

	End If
	Set objRs = Nothing
End Sub
%>