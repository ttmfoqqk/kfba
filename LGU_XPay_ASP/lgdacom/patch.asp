<%
    configPath = "C:/lgdacom"  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.

    '/*
    ' * [LG유플러스 환경파일 UPDATE]
    ' *
    ' * 이 페이지는 LG유플러스 환경파일을 UPDATE 합니다.(수정하지 마세요.)
    ' */ 
    CST_PLATFORM 	= trim(request("CST_PLATFORM"))         
    CST_MID 		= trim(request("CST_MID"))
    
    if CST_PLATFORM = "" then
    	Response.Write("[TX_PATCH error] CST_PLATFORM 파라미터 누락<br>")
        Response.End
    end if

    if CST_MID = "" then
    	Response.Write("[TX_PATCH error] CST_MID 파라미터 누락<br>")        
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
    
	Response.Write("patch result = " & xpay.Patch("lgdacom.conf") & "<P>")
%>
