<?xml version="1.0" encoding="euc-kr" ?>
<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%response.ContentType = "text/xml"%>
<%
Dim arrList
Dim cntList : cntList = -1

Dim DataMsg : DataMsg = "<data><admin_login>login</admin_login></data>"
Dim Idx     : Idx     = RequestSet("idx","POST",0)

If Session("UserIdx") <> "" Then
	Call Expires()
	Call dbopen()
		Call getViewCode()

		DataMsg = "<data>"
		DataMsg = DataMsg &  "<login>success</login>"
		If cntList > -1 Then 
			DataMsg = DataMsg &  "<Pay><![CDATA[" & FV_Pay & "]]></Pay>"
			DataMsg = DataMsg &  "<Payhtml><![CDATA[" & FormatNumber(FV_Pay,0) & "]]></Payhtml>"
			For iLoop = 0 To cntList
				DataMsg = DataMsg &  "<item>"
				DataMsg = DataMsg &  "<idx><![CDATA["   & arrList(FI_Idx,iLoop)    & "]]></idx>"
				DataMsg = DataMsg &  "<name><![CDATA["  & TagDecode( Trim( arrList(FI_Name,iLoop) ) )   & "]]></name>"
				DataMsg = DataMsg &  "<addr><![CDATA["  & TagDecode( Trim( arrList(FI_Addr,iLoop) ) )   & "]]></addr>"
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
	"DECLARE @Idx INT;" &_
	"SET @Idx = ?; " &_

	"SELECT " &_
	"	 ISNULL([Pay],0) AS [Pay] " &_
	"FROM [dbo].[SP_PROGRAM] WHERE [Idx] = @Idx " &_

	"SELECT " &_
	"	 A.[Idx] " &_
	"	,A.[Name] " &_
	"	,A.[Addr] " &_
	"FROM [dbo].[SP_PROGRAM_AREA] A " &_
	"INNER JOIN [dbo].[SP_PROGRAM_ON_AREA] B ON(A.[Idx] = B.[AreaIdx]) " &_
	"WHERE B.[ProgramIdx] = @Idx AND A.[Dellfg] = 0 AND ( A.[Code] IS NOT NULL OR A.[Code] !='' )  "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , Idx )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")

	'검정시행일 프로그램 목록
	set objRs = objRs.NextRecordset
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
	End If
	objRs.close	: Set objRs = Nothing
End Sub
%>