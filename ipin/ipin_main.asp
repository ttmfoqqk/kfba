<%
	'********************************************************************************************************************************************
	'NICE신용평가정보 Copyright(c) KOREA INFOMATION SERVICE INC. ALL RIGHTS RESERVED
	'
	'서비스명 : 가상주민번호서비스 (IPIN) 서비스
	'페이지명 : 가상주민번호서비스 (IPIN) 호출 페이지
    '********************************************************************************************************************************************
    
    DIM sSiteCode, sSitePw, sReturnURL, sCPRequest
    DIM IPIN_DLL, clsIPINDll
    DIM iRtn, sEncReqData, sRtnMsg
	Dim pageMode : pageMode = Request("pageMode")

	sSiteCode = ""			'IPIN 서비스 사이트 코드		(NICE신용평가정보에서 발급한 사이트코드)
	sSitePw   = ""		'IPIN 서비스 사이트 패스워드	(NICE신용평가정보에서 발급한 사이트패스워드)
	
	
	
	'┌ sReturnURL 변수에 대한 설명  ─────────────────────────────────────────────────────
	'	NICE신용평가정보 팝업에서 인증받은 사용자 정보를 암호화하여 귀사로 리턴합니다.
	'	따라서 암호화된 결과 데이타를 리턴받으실 URL 정의해 주세요.
	
	'	* URL 은 http 부터 입력해 주셔야하며, 외부에서도 접속이 유효한 정보여야 합니다.
	'	* 당사에서 배포해드린 샘플페이지 중, ipin_process.asp 페이지가 사용자 정보를 리턴받는 예제 페이지입니다.
	
	'	아래는 URL 예제이며, 귀사의 서비스 도메인과 서버에 업로드 된 샘플페이지 위치에 따라 경로를 설정하시기 바랍니다.
	'	예 - http://www.test.co.kr/ipin_process.asp, https://www.test.co.kr/ipin_process.asp, https://test.co.kr/ipin_process.asp
	'└────────────────────────────────────────────────────────────────────
	sReturnURL				= "http://" & Request.ServerVariables("SERVER_NAME") & "/ipin/ipin_result.asp"
	
	
	'┌ sCPRequest 변수에 대한 설명  ─────────────────────────────────────────────────────
	'	[CP 요청번호]로 귀사에서 데이타를 임의로 정의하거나, 당사에서 배포된 모듈로 데이타를 생성할 수 있습니다.
	
	'	CP 요청번호는 인증 완료 후, 암호화된 결과 데이타에 함께 제공되며
	'	데이타 위변조 방지 및 특정 사용자가 요청한 것임을 확인하기 위한 목적으로 이용하실 수 있습니다.
	
	'	따라서 귀사의 프로세스에 응용하여 이용할 수 있는 데이타이기에, 필수값은 아닙니다.
	'└────────────────────────────────────────────────────────────────────
	sCPRequest				= ""
	
	
	
	
	
	'IPIN 서비스 객체 정의
	IPIN_DLL		= "IPINClient.Kisinfo"
	
	'컴포넌트 객체 생성
	SET clsIPINDll	= Server.CreateObject(IPIN_DLL)
	
	
	'앞서 설명드린 바와같이, CP 요청번호는 배포된 모듈을 통해 아래와 같이 생성할 수 있습니다.
	clsIPINDll.fnRequestSEQ(sSiteCode)
	sCPRequest = clsIPINDll.bstrRandomRequestSEQ
	
	'CP 요청번호를 세션에 저장합니다.
	'현재 예제로 저장한 세션은 ipin_result.asp 페이지에서 데이타 위변조 방지를 위해 확인하기 위함입니다.
	'필수사항은 아니며, 보안을 위한 권고사항입니다.
	SESSION("CPREQUEST") = sCPRequest
	
	
	'Method 결과값(iRtn)에 따라, 프로세스 진행여부를 파악합니다.
	iRtn = clsIPINDll.fnRequest(sSiteCode, sSitePw, sCPRequest, sReturnURL)
	
	'Method 결과값에 따른 처리사항
	IF (iRtn = 0) THEN
	
		'fnRequest 함수 처리시 업체정보를 암호화한 데이터를 추출합니다.
		'추출된 암호화된 데이타는 당사 팝업 요청시, 함께 보내주셔야 합니다.
		sEncReqData = clsIPINDll.bstrRequestCipherData
		
		sRtnMsg = "정상 처리되었습니다."
	
	ELSEIF (iRtn = -9) THEN
    	sRtnMsg = "입력값 오류 : fnRequest 함수 처리시, 필요한 4개의 파라미터값의 정보를 정확하게 입력해 주시기 바랍니다."
    	sEncReqData = ""
    ELSE
    	sRtnMsg = "iRtn 값 확인 후, NICE신용평가정보 개발 담당자에게 문의해 주세요."
    	sEncReqData = ""
    END IF
    
    SET clsIPINDll = NOTHING

%>

<html>
<head>
	<title>NICE신용평가정보 가상주민번호 서비스</title>
	
	<script language='javascript'>
	window.onload = function(){
		fnPopup();
	};
	
	function fnPopup(){
		if('<%=iRtn%>' != '0'){alert('<%=sRtnMsg%>');self.close();return false;}
		document.form_ipin.action = "https://cert.vno.co.kr/ipin.cb";
		document.form_ipin.submit();
	}
	</script>
</head>

<body>
<!-- 가상주민번호 서비스 팝업을 호출하기 위해서는 다음과 같은 form이 필요합니다. -->
<form name="form_ipin" method="post">
	<input type="hidden" name="m" value="pubmain">						<!-- 필수 데이타로, 누락하시면 안됩니다. -->
    <input type="hidden" name="enc_data" value="<%= sEncReqData %>">	<!-- 위에서 업체정보를 암호화 한 데이타입니다. -->
    
    <!-- 업체에서 응답받기 원하는 데이타를 설정하기 위해 사용할 수 있으며, 인증결과 응답시 해당 값을 그대로 송신합니다.
    	 해당 파라미터는 추가하실 수 없습니다. -->
    <input type="hidden" name="param_r1" value="<%=pageMode%>">
    <input type="hidden" name="param_r2" value="">
    <input type="hidden" name="param_r3" value="">
</form>
</body>
</html>