<%
Dim objXMLHTTP : Set objXMLHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	objXMLHTTP.setOption 2, 13056
	objXMLHTTP.Open "GET", "http://kfbd.co.kr/LGU_XPay_ASP/lgdacom/conf/lgdacom.conf" , false
	objXMLHTTP.SetRequestHeader "Content-Type","application/x-www-form-urlencoded;"
	objXMLHTTP.SetRequestHeader "User-Agent", "Classic ASP VBScript OAuth"
	objXMLHTTP.Send()
	
	Response.write objXMLHTTP.ResponseText

Set objXMLHTTP = Nothing
On Error Goto 0
%>