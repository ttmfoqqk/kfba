<%
Dim BC_ARRY_LIST
Dim BC_CNT_LIST  : BC_CNT_LIST  = -1
Dim BC_FIRST_KEY : BC_FIRST_KEY = 0


'등록된 게시판 목록
Sub BoardCodeList()
	SET objRs  = Server.CreateObject("ADODB.RecordSet")
	SET objCmd = Server.CreateObject("adodb.command")

	SQL = "SELECT [Idx] , [Name] FROM [dbo].[SP_BOARD_CODE] WHERE [State] = 0 ORDER BY [Order] ASC,[Idx] DESC; "
	call cmdopen()
	with objCmd
		.CommandText = SQL
		Set objRs = .Execute
	End with
	call cmdclose()
	
	CALL setFieldIndex(objRs, "BDL")
	If NOT(objRs.BOF or objRs.EOF) Then
		BC_ARRY_LIST = objRs.GetRows()
		BC_CNT_LIST  = UBound(BC_ARRY_LIST, 2)
		BC_FIRST_KEY = BC_ARRY_LIST(BDL_Idx, 0)
	End If
	Set objRs = Nothing
End Sub
'등록된 게시판 상세설정
Sub BoardCodeView()
	SET objRs  = Server.CreateObject("ADODB.RecordSet")
	SET objCmd = Server.CreateObject("adodb.command")

	SQL = "SELECT "& vbCrLf &_
	"	 [Idx] "& vbCrLf &_
	"	,[Type] "& vbCrLf &_
	"	,[Name] "& vbCrLf &_
	"	,[PmsL] "& vbCrLf &_
	"	,[PmsV] "& vbCrLf &_
	"	,[PmsW] "& vbCrLf &_
	"	,[State] "& vbCrLf &_
	"	,[Order] "& vbCrLf &_
	"	,[Replyfg] "& vbCrLf &_
	"	,[CommentFg] "& vbCrLf &_
	"FROM [dbo].[SP_BOARD_CODE] "& vbCrLf &_
	"WHERE [Idx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput, 0 , BoardKey )
		Set objRs = .Execute		
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "BDV")
	Set objRs = Nothing

	If BDV_Idx = "" Then 
		Call msgbox("잘못된 게시판 정보 입니다.",true)
	End If
End Sub
%>