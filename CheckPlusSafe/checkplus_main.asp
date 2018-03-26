<%
Dim clsCPClient
Dim iRtn, sEncData, sPlainData
Dim sRequestNO, sSiteCode, sSitePassword , sReturnUrl , sErrorUrl, popgubun, customize

Dim pageMode : pageMode = Request("pageMode")

SET clsCPClient  = SERVER.CREATEOBJECT("CPClient.NiceID")

sSiteCode       = ""			'NICE로부터 부여받은 사이트 코드
sSitePassword   = ""			'NICE로부터 부여받은 사이트 패스워드
sAuthType       = "M"		'없으면 기본 선택화면, M: 핸드폰, C: 카드, X: 공인인증서
popgubun 	    = "N"		'Y : 취소버튼 있음 / N : 취소버튼 없음
customize       = ""		'없으면 기본 웹페이지 / Mobile : 모바일페이지
   
'CheckPlus(본인인증) 처리 후, 결과 데이타를 리턴 받기위해 다음예제와 같이 http부터 입력합니다.
sReturnUrl      = "http://" & Request.ServerVariables("SERVER_NAME") & "/CheckPlusSafe/checkplus_success.asp"			'성공시 이동될 URL
sErrorUrl	    = "http://" & Request.ServerVariables("SERVER_NAME") & "/CheckPlusSafe/checkplus_fail.asp"				'실패시 이동될 URL

sRequestNO = "REQ0000000001"			'요청 번호, 이는 성공/실패후에 같은 값으로 되돌려주게 되므로
								'업체에 적절하게 변경하여 쓰거나, 아래와 같이 생성한다.
iRtn = clsCPClient.fnRequestNO(sSiteCode)

IF iRtn = 0 THEN
	sRequestNO = clsCPClient.bstrRandomRequestNO
	session("REQ_SEQ") = sRequestNO		'해킹등의 방지를 위하여 세션을 사용한다면, 세션에 요청번호를 넣는다.
END IF

sPlainData = fnGenPlainData(sRequestNO, sSiteCode, sAuthType, sReturnUrl, sErrorUrl, popgubun, customize)

'실제적인 암호화
iRtn = clsCPClient.fnEncode(sSiteCode, sSitePassword, sPlainData)

IF iRtn = 0 THEN
	sEncData = clsCPClient.bstrCipherData
ELSE
	RESPONSE.WRITE "요청정보_암호화_오류:" & iRtn & "<br>"
	' -1 : 암호화 시스템 에러입니다.
	' -2 : 암호화 처리오류입니다.
	' -3 : 암호화 데이터 오류입니다.
	' -4 : 입력 데이터 오류입니다.
END IF

Set clsCPClient = Nothing
%>

<%
'**************************************************************************************
'문자열 생성 
'**************************************************************************************  					          	
Function fnGenPlainData(aRequestNO, aSiteCode, aAuthType, aReturnUrl, aErrorUrl, popgubun, customize)
			
	'입력 파라미터로 plaindata 생성 			
	retPlainData  = "7:REQ_SEQ" & fnGetDataLength(aRequestNO) & ":" & aRequestNO & _
								  "8:SITECODE" & fnGetDataLength(aSiteCode) & ":" & aSiteCode & _
								  "9:AUTH_TYPE" & fnGetDataLength(aAuthType) & ":" & aAuthType & _
								  "7:RTN_URL" & fnGetDataLength(aReturnUrl) & ":" & aReturnUrl & _
								  "7:ERR_URL" & fnGetDataLength(aErrorUrl) & ":" & aErrorUrl	& _	
								  "11:POPUP_GUBUN" & fnGetDataLength(popgubun) & ":" & popgubun & _
								  "9:CUSTOMIZE" & fnGetDataLength(customize) & ":" & customize
	fnGenPlainData = retPlainData		

End Function 

'**************************************************************************************
'입력파라미터의 문자열길이 추출	
'**************************************************************************************  					          	
Function fnGetDataLength(aData)		
	Dim iData_len
	if (len(aData) > 0) then
		for i = 1 to len(aData)
			if (ASC(mid(aData,i,1)) < 0) then	'한글인경우
				iData_len = iData_len + 2
			else			'한글이아닌경우
				iData_len = iData_len + 1
			end if		
		next
	else
		iData_len = 0
	end if
	
	fnGetDataLength = iData_len
End Function
%>

<html>
<head>
	<title>NICE신용평가정보 - CheckPlus 안심본인인증 테스트</title>
	
	<script language='javascript'>
	window.onload = function(){
		fnPopup();
	};
	
	function fnPopup(){
		if('<%=iRtn%>' != '0'){alert('err code : <%=iRtn%> 암호화 시스템 에러입니다. 잠시후에 시도해 주시기 바랍니다.');self.close();return false;}
		document.form_chk.action = "https://nice.checkplus.co.kr/CheckPlusSafeModel/checkplus.cb";
		document.form_chk.submit();
	}
	</script>
</head>
<body>
	<!-- 본인인증 서비스 팝업을 호출하기 위해서는 다음과 같은 form이 필요합니다. -->
	<form name="form_chk" method="post">
		<input type="hidden" name="m" value="checkplusSerivce">						<!-- 필수 데이타로, 누락하시면 안됩니다. -->
		<input type="hidden" name="EncodeData" value="<%= sEncData %>">		<!-- 위에서 업체정보를 암호화 한 데이타입니다. -->
	    
	    <!-- 업체에서 응답받기 원하는 데이타를 설정하기 위해 사용할 수 있으며, 인증결과 응답시 해당 값을 그대로 송신합니다.
	    	   해당 파라미터는 추가하실 수 없습니다. -->
		<input type="hidden" name="param_r1" value="<%=pageMode%>">
		<input type="hidden" name="param_r2" value="">
		<input type="hidden" name="param_r3" value="">
	</form>
</body>
</html>