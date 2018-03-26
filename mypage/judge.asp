<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/template.class.asp" -->
<!-- #include file = "../_lib/board.sub.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim arrList
Dim cntList  : cntList  = -1
Dim cntTotal : cntTotal = 0
Dim rows     : rows     = 10
Dim pageNo   : pageNo   = RequestSet("pageNo"   ,"GET",1 )

Dim PageParams : PageParams = "pageNo=" & pageNo
Dim pageUrl    : pageUrl    = g_url & "?" & "pageNo=__PAGE__"

checkLogin( g_host & g_url &"?"&PageParams )

Call Expires()
Call dbopen()
	Call getList()
	Call CheckApplicationCnt()
Call dbclose()

Dim pagelist : pagelist = printPageList(cntTotal, pageNo, rows, pageUrl)

dim ntpl
set ntpl = new SkyTemplate
ntpl.setBlockErrorCheck(false)
ntpl.setTplDir( FRONT_ROOT_DIR & TPL_DIR_FOLDER )

ntpl.setFile array( _
	 array("HEADER" , "_inc/header.html") _
	,array("MAIN"   , "mypage/judge.html" ) _
	,array("FOOTER" , "_inc/footer.html") _
), ""
'// 상단 로그인 블럭처리
Call loginBlock_ntpl("HEADER","LOGIN_BOX","LOGOUT_BOX")

call ntpl.setBlock("MAIN", array("BOARD_LOOP" , "BOARD_LOOP_NODATA" , "MYJUDGE_LIST"))

'왼쪽 심사위원등록 메뉴
If LEFT_JUDGE_MENU_CNT > 0 Then 
	ntpl.tplParseBlock("MYJUDGE_LIST")
Else
	ntpl.tplBlockDel("MYJUDGE_LIST")
End If 


If cntList > -1 Then 

	for iLoop = 0 to cntList

	StateTxt = ""

	If arrList(FI_State,iLoop) = "0" Then 
		StateTxt = "<font color=""#11179a"">승인</font>"
	ElseIf arrList(FI_State,iLoop) = "1" Then 
		StateTxt = "접수"
	ElseIf arrList(FI_State,iLoop) = "2" Then 
		StateTxt = "<font color=""#9a1134"">불합격</font>"
	End If

		ntpl.setBlockReplace array( _
			 array("Number", arrList(FI_rownum,iLoop) ) _
			,array("Idx", arrList(FI_Idx,iLoop) ) _
			,array("ProgramNema", arrList(FI_ProgramNema,iLoop) & IIF(arrList(FI_ProgramKind,iLoop)="1"," [필기]"," [실기]") ) _
			,array("OnData", arrList(FI_OnData,iLoop) ) _
			,array("AreaName", arrList(FI_AreaName,iLoop) ) _
			,array("State", StateTxt ) _
			,array("InDate", arrList(FI_InData,iLoop) ) _
		), ""
		ntpl.tplParseBlock("BOARD_LOOP")
	Next 
	ntpl.tplBlockDel("BOARD_LOOP_NODATA")
Else
	ntpl.tplParseBlock("BOARD_LOOP_NODATA")
	ntpl.tplBlockDel("BOARD_LOOP")
End If

ntpl.tplAssign array(   _
	 array("imgDir"   , TPL_DIR_IMAGES ) _
	,array("pageNo"   , pageNo ) _
	,array("pagelist" , pagelist ) _
	,array("PageParams" , PageParams ) _
), ""

'// 예제에서 { 마크 사용을 위한 것
ntpl.tplAssign "m", "{"

ntpl.tplParse()  '// 설정한 템플릿 파일처리
ntpl.tplPrint()  '// 출력

set ntpl = Nothing






Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON;" &_
	"DECLARE @pageNo INT, @rows INT, @UserIdx INT ;" &_
	"SET @pageNo = ?; SET @rows = ?; SET @UserIdx = ? ;" &_

	"WITH LIST AS " &_
	"( " &_
	"	SELECT row_number() over (order by [Idx] asc ) as [rownum] " &_
	"		, count(*) over () as [tcount] " &_
	"		,[Idx] " &_
	"		,[State] " &_
	"		,CONVERT(VARCHAR,[InData],23) AS [InData] " &_
	"		,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = [ProgramIdx] ) AS [ProgramNema]" &_
	"		,[ProgramKind] " &_
	"	FROM [dbo].[SP_PROGRAM_JUDGE_APP] " &_
	"	WHERE [DellFg] = 0 AND [UserIdx] = @UserIdx " &_
	") " &_
	"SELECT L.* " &_
	"FROM LIST L " &_
	"WHERE (tcount-rownum+1) BETWEEN dbo.fnc_row_fr(@pageNo,@rows,tcount) AND dbo.fnc_row_to(@pageNo,@rows,tcount) " &_
	"ORDER BY rownum desc "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		
		.Parameters.Append .CreateParameter( "@pageNo"    ,adInteger , adParamInput ,  0  , pageNo )
		.Parameters.Append .CreateParameter( "@rows"      ,adInteger , adParamInput ,  0  , rows )
		.Parameters.Append .CreateParameter( "@UserIdx"   ,adInteger , adParamInput ,  0  , Session("UserIdx") )

		set objRs = .Execute
	End with
	call cmdclose()

	CALL setFieldIndex(objRs, "FI")
	If Not(objRs.Eof or objRs.Bof) Then		
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
		cntTotal	= arrList(FI_tcount, 0)
	End If
	Set objRs = Nothing
End Sub
%>