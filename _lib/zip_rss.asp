<?xml version="1.0" encoding="utf-8" ?>
<!-- #include file = "charSetutf8.asp" -->
<!-- #include file = "common.asp" -->
<%response.ContentType = "text/xml"%>
<%
Dim arrList,DataMsg
Dim cntList : cntList = -1
Dim schDong : schDong = Trim( Request("schDong") )

Call Expires()
Call dbopen()
	Call getList()

	DataMsg = "<data>"
	For iLoop = 0 To cntList
		DataMsg = DataMsg &  "<item>"
		DataMsg = DataMsg &  "<zipcode><![CDATA[" & arrList(FI_zipcode,iLoop) & "]]></zipcode>"
		DataMsg = DataMsg &  "<addr><![CDATA["    & arrList(FI_addr,iLoop)    & "]]></addr>"
		DataMsg = DataMsg &  "<bunji><![CDATA["   & arrList(FI_bunji,iLoop)   & "]]></bunji>"
		DataMsg = DataMsg &  "</item>"
	Next
	DataMsg = DataMsg &  "</data>"

Call dbclose()


Response.write DataMsg

Sub getList()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("adodb.command")

	SQL = "SELECT DISTINCT zipcode " &_
	"	, sido + ' ' + gugun + ' ' + dong AS addr " &_
	"	, ISNULL(bunji,'') AS bunji " &_
	"FROM [dbo].[SP_ZIPCODE] " &_
	"WHERE dong like '%'+?+'%' " &_
	"ORDER BY zipcode asc "
	
	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@schDong", adVarChar, adParamInput, 50, IIF(LEN(schDong)>0, schDong, "NO-CONDITION") )
		Set objRs = .Execute
	End with
	set objCmd = nothing
	CALL setFieldIndex(objRs, "FI")
	If Not(objRs.Eof or objRs.Bof) Then
		arrList = objRs.GetRows()
		cntList = UBound(arrList, 2)
	End If
	objRs.close	: Set objRs = Nothing
End Sub
%>