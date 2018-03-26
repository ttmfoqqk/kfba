<%
    '/*
    ' * [현금영수증 발급 요청 페이지]
    ' *
    ' * 파라미터 전달시 POST를 사용하세요
    ' */
    CST_PLATFORM               = trim(request("CST_PLATFORM"))       	'LG유플러스 결제 서비스 선택(test:테스트, service:서비스)
    CST_MID                    = trim(request("CST_MID"))            	'상점아이디(LG유플러스으로 부터 발급받으신 상점아이디를 입력하세요)
                                                                     	'테스트 아이디는 't'를 반드시 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                                    	'상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_METHOD                 = trim(request("LGD_METHOD"))         	 '메소드('AUTH':승인, 'CANCEL' 취소)
    LGD_OID                    = trim(request("LGD_OID"))            	 '주문번호(상점정의 유니크한 주문번호를 입력하세요)
	LGD_PAYTYPE                = trim(request("LGD_PAYTYPE"))         	 '결제수단 코드 (SC0030:계좌이체, SC0040:가상계좌, SC0100:무통장입금 단독)
    LGD_AMOUNT                 = trim(request("LGD_AMOUNT"))         	 '금액("," 를 제외한 금액을 입력하세요)
    LGD_CASHCARDNUM            = trim(request("LGD_CASHCARDNUM"))    	 '발급번호(주민등록번호,현금영수증카드번호,휴대폰번호 등등)
    LGD_CUSTOM_MERTMAME        = trim(request("LGD_CUSTOM_MERTNAME"))    '상점명
    LGD_CUSTOM_BUSINESSNUM     = trim(request("LGD_CUSTOM_BUSINESSNUM")) '사업자등록번호
    LGD_CUSTOM_MERTPHONE       = trim(request("LGD_CUSTOM_MERTPHONE")) 	 '상점 전화번호
    LGD_CASHRECEIPTUSE     	   = trim(request("LGD_CASHRECEIPTUSE")) 	 '현금영수증발급용도('1':소득공제, '2':지출증빙)
    LGD_PRODUCTINFO     	   = trim(request("LGD_PRODUCTINFO")) 	     '상품명
    LGD_TID     	   		   = trim(request("LGD_TID")) 	 		     'LG유플러스 거래번호
    
    configPath				   = "C:/lgdacom"  							 'LG유플러스에서 제공한 환경파일의 위치
    
	Dim xpay
	Dim i, j
	Dim itemName
	
	Set xpay = server.CreateObject("XPayClientCOM.XPayClient")	
    xpay.Init configPath, CST_PLATFORM    
    xpay.Init_TX(LGD_MID)
    xpay.Set "LGD_TXNAME", "CashReceipt"
	xpay.Set "LGD_METHOD", LGD_METHOD
	xpay.Set "LGD_PAYTYPE", LGD_PAYTYPE
	
	if LGD_METHOD = "AUTH" then              '현금영수증 발급 요청 
		xpay.Set "LGD_OID", LGD_OID 
		xpay.Set "LGD_AMOUNT", LGD_AMOUNT
		xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM
		xpay.Set "LGD_CUSTOM_MERTMAME", LGD_CUSTOM_MERTMAME
		xpay.Set "LGD_CUSTOM_BUSINESSNUM", LGD_CUSTOM_BUSINESSNUM
		xpay.Set "LGD_CUSTOM_MERTPHONE", LGD_CUSTOM_MERTPHONE
		xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
		
		if LGD_PAYTYPE = "SC0030" then		'기결제된 계좌이체건 현금영수증 발급요청시 필수 
			xpay.Set "LGD_TID", LGD_TID
		elseIf	LGD_PAYTYPE = "SC0040" then	'기결제된 가상계좌건 현금영수증 발급요청시 필수  
			xpay.Set "LGD_TID", LGD_TID
			xpay.Set "LGD_SEQNO", "001"
		else								'무통장입금 단독건 발급요청  
			xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
		end if
	
	else									'현금영수증 취소 요청
		xpay.Set "LGD_TID", LGD_TID
		
		if	LGD_PAYTYPE = "SC0040" then		'가상계좌건 현금영수증 발급취소시 필수
			xpay.Set "LGD_SEQNO", "001"
		end if	
	end if
 
    xpay.TX()	
	Response.Write("현금영수증 처리가 완료되었습니다. <br>")	
	Response.Write("TX Response_code = " & xpay.resCode & "<br>")
	Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")	
    
    Response.Write("결과코드 : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
    Response.Write("결과메세지 : " & xpay.Response("LGD_RESPMSG", 0) & "<p>")

    Response.Write("[결과 파라미터]<br>")

   '아래는 결과 파라미터를 모두 찍어 줍니다.   
	Dim itemCount
    Dim resCount
    itemCount = xpay.resNameCount
    resCount = xpay.resCount
	
    For i = 0 To itemCount - 1
        itemName = xpay.ResponseName(i)
		Response.Write(itemName & "&nbsp:&nbsp")
        For j = 0 To resCount - 1
			Response.Write(xpay.Response(itemName, j) & "<br>")
        Next
    Next	
%>
