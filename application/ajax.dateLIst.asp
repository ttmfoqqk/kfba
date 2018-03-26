<?xml version="1.0" encoding="euc-kr" ?>
<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%response.ContentType = "text/xml"%>
<%
Dim arrList
Dim cntList : cntList = -1

Dim DataMsg : DataMsg = "<data><login>login</login></data>"
Dim Idx     : Idx     = RequestSet("idx" ,"POST",0)
Dim Kind    : Kind    = RequestSet("Kind","POST",0)
Dim Pclass  : Pclass  = RequestSet("Class","POST",0)

If Session("UserIdx") <> "" Then
	Call Expires()
	Call dbopen()
		Call getViewCode()

		DataMsg = "<data>"
		DataMsg = DataMsg &  "<login>success</login>"
		DataMsg = DataMsg &  "<check><![CDATA["   & CHECK_CK_Kind & "]]></check>"
		If cntList > -1 Then 
			For iLoop = 0 To cntList
				DataMsg = DataMsg &  "<item>"
				DataMsg = DataMsg &  "<idx><![CDATA["   & arrList(FI_Idx,iLoop) & "]]></idx>"
				DataMsg = DataMsg &  "<date><![CDATA["  & arrList(FI_OnData,iLoop) & "]]></date>"
				DataMsg = DataMsg &  "</item>"
			Next
		End If
		DataMsg = DataMsg &  "</data>"
	Call dbclose()
End If

Response.write DataMsg

Sub getViewCode()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Idx INT,@Kind INT,@Class INT,@CK_Kind INT,@UserIdx INT;" &_
	"SET @Idx     = ?;" &_
	"SET @Kind    = ?;" &_
	"SET @Class   = ?;" &_
	"SET @UserIdx = ?;" &_
	"SET @CK_Kind = 1;" &_
	
	"IF @Kind >= 2 " &_
	"BEGIN " &_
	"	SET @CK_Kind = ( " &_
	"		SELECT COUNT(*) " &_
	"		FROM [dbo].[SP_PROGRAM] A " &_
	"		INNER JOIN [dbo].[SP_PROGRAM_APP] B ON(A.[Idx] = B.[ProgramIdx]) " &_
	"		WHERE A.[CodeIdx] = @Idx " &_
	"		AND [UserIdx] = @UserIdx " &_
	"		AND [Kind]  = 1 " &_
	"		AND [Class] = @Class " &_
	"		AND [State] = 10 " &_
	"	)" &_
	"END " &_

	"SELECT " &_
	"	 [Idx] " &_
	"	,[OnData] " &_
	"FROM [dbo].[SP_PROGRAM] A " &_
	"LEFT JOIN ( " &_
	"	SELECT " &_
	"		 [ProgramIdx] " &_
	"		,COUNT(*) AS [CNT_APP] " &_
	"	FROM [dbo].[SP_PROGRAM_APP] " &_
	"	WHERE [State] != 2 " &_
	"	GROUP BY [ProgramIdx] " &_
	") B ON(A.[Idx] = B.[ProgramIdx] ) " &_
	"WHERE [CodeIdx] = @Idx " &_
	"AND [Dellfg] = 0 " &_ 
	"AND [Kind] = @Kind " &_ 
	"AND [Class] = @Class " &_ 
	"AND CONVERT(varchar(10),[StartDate],23) <= CONVERT(varchar(10),getDate(),23) " &_
	"AND CONVERT(varchar(10),[EndDate],23) >= CONVERT(varchar(10),getDate(),23) " &_
	"AND A.[MaxNumber] > ISNULL(B.[CNT_APP],0) " &_
	"ORDER BY [OnData] ASC " &_

	"SELECT @CK_Kind AS [CK_Kind] "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"     ,adInteger , adParamInput , 0 , Idx )
		.Parameters.Append .CreateParameter( "@Kind"    ,adInteger , adParamInput , 0 , Kind )
		.Parameters.Append .CreateParameter( "@Class"   ,adInteger , adParamInput , 0 , Pclass )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()

	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
	End If

	'실기전 필기 검사
	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "CHECK")

	objRs.close	: Set objRs = Nothing
End Sub
%>