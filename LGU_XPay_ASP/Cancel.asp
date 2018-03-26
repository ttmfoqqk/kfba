<%
    '/*
    ' * [결제취소 요청 페이지]
    ' *
    ' * LG유플러스으로 부터 내려받은 거래번호(LGD_TID)를 가지고 취소 요청을 합니다.(파라미터 전달시 POST를 사용하세요)
    ' * (승인시 LG유플러스으로 부터 내려받은 PAYKEY와 혼동하지 마세요.)
    ' */

    CST_PLATFORM         = trim(request("CST_PLATFORM"))        ' LG유플러스 결제서비스 선택(test:테스트, service:서비스)
    CST_MID              = trim(request("CST_MID"))             ' LG유플러스으로 부터 발급받으신 상점아이디를 입력하세요.
                                                                ' 테스트 아이디는 't'를 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                               ' 상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_TID                 = trim(request("LGD_TID"))                  ' LG유플러스으로 부터 내려받은 거래번호(LGD_TID)
    LGD_CANCELREASON        = trim(request("LGD_CANCELREASON"))         ' 취소사유
    LGD_CANCELREQUESTER     = trim(request("LGD_CANCELREQUESTER"))      ' 취소요청자
    LGD_CANCELREQUESTERIP   = trim(request("LGD_CANCELREQUESTERIP"))    ' 취소요청IP
    

    configPath           = "C:/lgdacom"         				' LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.


    Set xpay = CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID
    xpay.Set "LGD_CANCELREASON", LGD_CANCELREASON
    xpay.Set "LGD_CANCELREQUESTER", LGD_CANCELREQUESTER
    xpay.Set "LGD_CANCELREQUESTERIP", LGD_CANCELREQUESTERIP
 

    '/*
    ' * 1. 결제취소 요청 결과처리
    ' *
    ' * 취소결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
	' *
	' * [[[중요]]] 고객사에서 정상취소 처리해야할 응답코드
	' * 1. 신용카드 : 0000, AV11  
	' * 2. 계좌이체 : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (환불진행중 응답-> 환불결과코드.xls 참고)
	' * 3. 나머지 결제수단의 경우 0000(성공) 만 취소성공 처리
	' *
    ' */

    if xpay.TX() then
        '1)결제취소결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
        Response.Write("결제취소 요청이 완료되었습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    else
        '2)API 요청 실패 화면처리
        Response.Write("결제취소 요청이 실패하였습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if
%>
