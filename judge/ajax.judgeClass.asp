<?xml version="1.0" encoding="euc-kr" ?>
<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%response.ContentType = "text/xml"%>
<%
Dim DataMsg : DataMsg = "<data><login>login</login></data>"
Dim Idx     : Idx     = RequestSet("idx","POST",0)

If Session("UserIdx") <> "" Then
	Call Expires()
	Call dbopen()
		Call getViewCode()

		DataMsg = "<data>"
		DataMsg = DataMsg &  "<login>success</login>"
		DataMsg = DataMsg &  "<idx1><![CDATA["   & FV1_Idx & "]]></idx1>"
		DataMsg = DataMsg &  "<idx2><![CDATA["   & FV2_Idx & "]]></idx2>"
		DataMsg = DataMsg &  "<idx3><![CDATA["   & FV3_Idx & "]]></idx3>"
		DataMsg = DataMsg &  "</data>"
	Call dbclose()
End If

Response.write DataMsg

Sub getViewCode()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Idx INT,@UserIdx INT;" &_
	"SET @Idx     = ?; " &_
	"SET @UserIdx = ?; " &_

	"SELECT " &_
	"	 [Idx] " &_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP] " &_ 
	"WHERE [ProgramIdx] = @Idx " &_
	"AND [Dellfg] = 0 " &_
	"AND [State] != 2 " &_
	"AND [UserIdx] = @UserIdx " &_
	"AND [ProgramKind] = 1 " &_

	"SELECT " &_
	"	 [Idx] " &_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP] " &_ 
	"WHERE [ProgramIdx] = @Idx " &_
	"AND [Dellfg] = 0 " &_
	"AND [State] != 2 " &_
	"AND [UserIdx] = @UserIdx " &_
	"AND [ProgramKind] = 2 " &_

	"SELECT " &_
	"	 [Idx] " &_
	"FROM [dbo].[SP_PROGRAM_JUDGE_APP] " &_ 
	"WHERE [ProgramIdx] = @Idx " &_
	"AND [Dellfg] = 0 " &_
	"AND [State] != 2 " &_
	"AND [UserIdx] = @UserIdx " &_
	"AND [ProgramKind] = 3 "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx"     ,adInteger , adParamInput , 0 , Idx )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()
	'필기코드
	CALL setFieldValue(objRs, "FV1")

	'실기코드
	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "FV2")

	'SPECIAL코드
	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "FV3")

	objRs.close	: Set objRs = Nothing
End Sub
%>