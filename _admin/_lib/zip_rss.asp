<?xml version="1.0" encoding="utf-8" ?>
<!-- #include file = "../common/carset_utf8.asp" -->
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
	with objCmd
		.ActiveConnection  = objConn
		.prepared          = true
		.CommandType       = adCmdStoredProc
		.CommandText       = "OAPI_ZIPCODE_L"
		.Parameters("@schDong").value = IIF(LEN(schDong)>0, schDong, "NO-CONDITION")
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