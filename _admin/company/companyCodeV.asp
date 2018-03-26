<?xml version="1.0" encoding="euc-kr" ?>
<!-- #include file = "../../_lib/header.asp" -->
<!-- #include file = "../_lib/common.asp" -->
<%response.ContentType = "text/xml"%>
<%
Dim arrList
Dim cntList : cntList = -1

Dim DataMsg : DataMsg = "<data><admin_login>login</admin_login></data>"
Dim mode    : mode    = Trim( Request.Form("mode") )
Dim Idx     : Idx     = Trim( Request.Form("Idx") )

If Session("Admin_Idx") <> "" Then
	Call Expires()
	Call dbopen()
		If mode = "1" Then 
			Call getViewCode1()
		Else
			If isNumeric(Idx) Then 
				Call getViewCode2()
			End If
		End If
		DataMsg = "<data>"
		DataMsg = DataMsg &  "<admin_login>success</admin_login>"
		For iLoop = 0 To cntList
			DataMsg = DataMsg &  "<item>"
			DataMsg = DataMsg &  "<code_idx><![CDATA["   & arrList(FI_Idx,iLoop)    & "]]></code_idx>"
			DataMsg = DataMsg &  "<code_name><![CDATA["  & TagDecode( Trim( arrList(FI_Name,iLoop) ) )   & "]]></code_name>"
			DataMsg = DataMsg &  "<code_order><![CDATA[" & arrList(FI_Order,iLoop)  & "]]></code_order>"
			DataMsg = DataMsg &  "<code_bigo><![CDATA["  & TagDecode( Trim( arrList(FI_Bigo,iLoop) ) )   & "]]></code_bigo>"
			DataMsg = DataMsg &  "<code_usfg><![CDATA["  & arrList(FI_UsFg,iLoop) & "]]></code_usfg>"
			DataMsg = DataMsg &  "</item>"
		Next
		DataMsg = DataMsg &  "</data>"
	Call dbclose()
End If

Response.write DataMsg

Sub getViewCode1()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT " &_
	"	 [Idx] " &_
	"	,[Name] " &_
	"	,[Order] " &_
	"	,[Bigo] " &_
	"	,[UsFg] " &_
	"FROM [dbo].[SP_COMM_CODE1] " &_
	"WHERE [Idx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput,  0, Idx  )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
	End If
	objRs.close	: Set objRs = Nothing
End Sub

Sub getViewCode2()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT " &_
	"	 [Idx] " &_
	"	,[Name] " &_
	"	,[Order] " &_
	"	,[Bigo] " &_
	"	,[UsFg] " &_
	"FROM [dbo].[SP_COMM_CODE2] " &_
	"WHERE [Idx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput,  0, Idx  )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldIndex(objRs, "FI")
	If NOT(objRs.BOF or objRs.EOF) Then
		arrList		= objRs.GetRows()
		cntList		= UBound(arrList, 2)
	End If
	objRs.close	: Set objRs = Nothing
End Sub
%>