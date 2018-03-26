<!-- #include file = "../_lib/header.asp" -->
<%
Dim arrList
Dim cntList   : cntList  = -1

Dim startTs   : startTs   = Request("start")
Dim endTs     : endTs     = Request("end")
Dim startDate : startDate = DateAdd("s",startTs, CDate("1970-01-01 09:00:00"))
Dim endDate   : endDate   = DateAdd("s",endTs, CDate("1970-01-01 09:00:00"))

Call Expires()
Call dbopen()
	Call getList()
Call dbclose()

Dim html : html = "["
If cntList > -1 Then 
	
	Dim colon : colon = ","
	for iLoop = 0 to cntList
		If iLoop = cntList Then 
			colon = ""
		End If

		PrograName = arrList(FI_ProgramName,iLoop)

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

		ProgramLink = "write.asp?applicationKey=" & arrList(FI_CodeIdx,iLoop)
		' 마감
		If arrList(FI_EndDate,iLoop) < Left(Now(),10) Then
			ProgramLink = "javascript:void(alert('응시 마감되었습니다.'))"
		End If
		' 접수전
		If arrList(FI_StartDate,iLoop) > Left(Now(),10) Then 
			ProgramLink = "javascript:void(alert('응시 접수기간이 아닙니다.'))"
		End If
		' 인원제한
		If arrList(FI_MaxNumber,iLoop) <= arrList(FI_CNT_APP,iLoop) Then 
			ProgramLink = "javascript:void(alert('응시 정원초과!'))"
		End If


		html = html & "" &_
		"    {"&_
		"        ""title"":"""& PrograName &""","&_
		"        ""start"":"""& arrList(FI_OnData,iLoop) & ""","&_
		"        ""url"":""" & ProgramLink & ""","&_
		"        ""allDay"":false"&_
		"    }" & colon
	Next

End If
html = html & "]"
Response.write html


Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Sdate VARCHAR(10) , @Edate VARCHAR(10) ,@sPcode VARCHAR(10) ;" &_
	"SET @Sdate  = ?; " &_
	"SET @Edate  = ?; " &_
	"SET @sPcode = ?; " &_

	"WITH LIST AS( " &_
	"	SELECT row_number() over (order by [OnData] , [Idx] desc ) as [rownum]" &_
	"		,count(*) over () as [tcount] " &_
	"		,[Idx] " &_
	"		,[CodeIdx] " &_
	"		,convert(varchar,[OnData],20) AS [OnData] " &_
	"		,ISNULL( [Pay] , 0 ) AS [Pay] " &_
	"		,CONVERT(varchar(10),[StartDate],23) AS [StartDate] " &_
	"		,CONVERT(varchar(10),[EndDate],23) AS [EndDate] " &_
	"		,ISNULL( [MaxNumber] , 0 ) AS [MaxNumber] " &_
	"		,[Kind] " &_
	"		,[Class] " &_
	"		,( SELECT [Name] FROM [dbo].[SP_COMM_CODE2] WHERE [Idx] = A.[CodeIdx] ) AS [ProgramName] " &_
	"		,ISNULL(B.[CNT_APP],0) AS [CNT_APP] " &_
	"	FROM [dbo].[SP_PROGRAM] A " &_
	"	LEFT JOIN ( " &_
	"		SELECT " &_
	"			 [ProgramIdx] " &_
	"			,COUNT(*) AS [CNT_APP] " &_
	"		FROM [dbo].[SP_PROGRAM_APP] " &_
	"		WHERE [State] != 2 " &_
	"		GROUP BY [ProgramIdx] " &_
	"	) B ON(A.[Idx] = B.[ProgramIdx] ) " &_

	"   WHERE [Dellfg] = 0 " &_
	"   AND CASE @sPcode WHEN '' THEN '' ELSE [CodeIdx] END = @sPcode " &_
	") SELECT L.* " &_
	"FROM LIST L " &_
	"ORDER BY rownum desc; "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Sdate"  ,adVarChar , adParamInput , 10 , Left(startDate,10) )
		.Parameters.Append .CreateParameter( "@Edate"  ,adVarChar , adParamInput , 10 , Left(endDate,10) )
		.Parameters.Append .CreateParameter( "@sPcode" ,adVarChar , adParamInput , 20 , IIF(sPcode="","",sPcode) )
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
%>