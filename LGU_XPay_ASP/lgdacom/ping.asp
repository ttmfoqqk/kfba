<%
	configPath = ""  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.
 
    '/*
    ' * [LG유플러스 연결 확인 페이지]
    ' *
    ' * 이페이지는 LG유플러스과의 연결을 테스트 하는 페이지 입니다.(수정하지 마세요.)
    ' */
    'CST_PLATFORM 	= trim(request("CST_PLATFORM"))         
    'CST_MID 		= trim(request("CST_MID"))  't 가 추가되지 않은 가입요청시 아이디를 입력바랍니다.
    
	CST_PLATFORM 	= "test"
    CST_MID 		= ""
            
    if CST_PLATFORM = "" then
    	Response.Write("[TX_PING error] CST_PLATFORM 파라미터 누락<br>")
        Response.End
    end if

    if CST_MID = "" then
    	Response.Write("[TX_PING error] CST_MID 파라미터 누락<br>")        
        Response.End
    end if
    
    if CST_PLATFORM = "test" then
        LGD_MID = "t" & CST_MID                                   
    else
        LGD_MID = CST_MID                                         
    end if     
    
	Dim xpay
	Dim i, j
	Dim itemName
	
	Set xpay = server.CreateObject("XPayClientCOM.XPayClient")	
    xpay.Init configPath, CST_PLATFORM   
    xpay.Init_TX(LGD_MID)
    xpay.Set "LGD_TXNAME", "Ping"
    xpay.Set "LGD_RESULTCNT", "3"
    xpay.TX()	
	
	Response.Write("ResCode = " & xpay.resCode & "<P>")
	Response.Write("ResMsg = " & xpay.resMsg & "<P>")

	Response.Write("Response <P>")

    Dim itemCount
    Dim resCount
    itemCount = xpay.resNameCount
    resCount = xpay.resCount
	
    For i = 0 To itemCount - 1
        itemName = xpay.ResponseName(i)
		Response.Write(itemName & "<P>")
        For j = 0 To resCount - 1
			Response.Write(xpay.Response(itemName, j) & "<P>")
        Next
    Next
%>

<%=server.mapPath("/LGU_XPay_ASP/lgdacom/")%>