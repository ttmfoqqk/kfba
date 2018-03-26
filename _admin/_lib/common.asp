<%
' 게시판 다운로드 패치
Dim DOWNLOAD_BASE_PATH : DOWNLOAD_BASE_PATH = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/board/&file="
' 구직게시판 다운로드 패치
Dim DOWNLOAD_BASE_PATH_JOB : DOWNLOAD_BASE_PATH_JOB = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/job/&file="
' 개인 사진 패치
Dim USER_PHOTO_PATH : USER_PHOTO_PATH = FRONT_ROOT_DIR & "upload/appMember/"
Dim DOWNLOAD_USER_PHOTO_PATH : DOWNLOAD_USER_PHOTO_PATH = FRONT_ROOT_DIR & "_lib/dowload.asp?pach=/upload/appMember/&file="
'------------------------------------------------------------------------------------
'' 스킨 경로
'------------------------------------------------------------------------------------
Const TPL_DIR_FOLDER = "_skin/basic"
Const TPL_DIR_IMAGES = "../_skin/basic/images"
'------------------------------------------------------------------------------------
'' 관리자 로그인 체크.
'------------------------------------------------------------------------------------
Function checkAdminLogin(url)
	If session("Admin_Id")="" or IsNull(session("Admin_Id"))=True Then 
		response.redirect "../index.asp?GoUrl=" & server.urlencode(url)
	End If
End Function

Sub topMenuOver()
	If INSTR(LCase(g_url),"/company/")>0 Then
		ntpl.tplAssign array(   _
			 array("TopMenuOverClass1" , "top_menu_item_over" ) _
			,array("TopMenuOverClass2" , "" ) _
			,array("TopMenuOverClass3" , "" ) _
			,array("TopMenuOverClass4" , "" ) _
			,array("TopMenuOverClass5" , "" ) _
		), ""
	elseIf INSTR(LCase(g_url),"/programs/")>0 Then 
		ntpl.tplAssign array(   _
			 array("TopMenuOverClass1" , "" ) _
			,array("TopMenuOverClass2" , "top_menu_item_over" ) _
			,array("TopMenuOverClass3" , "" ) _
			,array("TopMenuOverClass4" , "" ) _
			,array("TopMenuOverClass5" , "" ) _
		), ""
	elseIf INSTR(LCase(g_url),"/application/")>0 Then 
		ntpl.tplAssign array(   _
			 array("TopMenuOverClass1" , "" ) _
			,array("TopMenuOverClass2" , "" ) _
			,array("TopMenuOverClass3" , "top_menu_item_over" ) _
			,array("TopMenuOverClass4" , "" ) _
			,array("TopMenuOverClass5" , "" ) _
		), ""
	elseIf INSTR(LCase(g_url),"/member/")>0 Then 
		ntpl.tplAssign array(   _
			 array("TopMenuOverClass1" , "" ) _
			,array("TopMenuOverClass2" , "" ) _
			,array("TopMenuOverClass3" , "" ) _
			,array("TopMenuOverClass4" , "top_menu_item_over" ) _
			,array("TopMenuOverClass5" , "" ) _
		), ""
	elseIf INSTR(LCase(g_url),"/customer/")>0 Then 
		ntpl.tplAssign array(   _
			 array("TopMenuOverClass1" , "" ) _
			,array("TopMenuOverClass2" , "" ) _
			,array("TopMenuOverClass3" , "" ) _
			,array("TopMenuOverClass4" , "" ) _
			,array("TopMenuOverClass5" , "top_menu_item_over" ) _
		), ""
	End If

End Sub
%>