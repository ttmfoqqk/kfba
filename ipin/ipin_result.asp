<!-- #include file = "../_lib/header.asp" -->
<%
'********************************************************************************************************************************************
'NICE신용평가정보 Copyright(c) KOREA INFOMATION SERVICE INC. ALL RIGHTS RESERVED
'
'서비스명 : 가상주민번호서비스 (IPIN) 서비스
'페이지명 : 가상주민번호서비스 (IPIN) 결과 페이지
'********************************************************************************************************************************************

DIM sSiteCode, sSitePw, sResponseData, sCPRequest
DIM IPIN_DLL, clsIPINDll
DIM iRtn, sRtnMsg
DIM sVNumber, sName, sDupInfo, sAgeCode, sGenderCode, sBirthDate, sNationalInfo, sCPRequestNum
Dim sCoInfo1, sCIUpdate ,sAuthInfo
Dim GoUrl : GoUrl = "joinData.asp"

Dim NoJoinDate  : NoJoinDate    = 0 ' 탈퇴후 재가입 가능한 날짜 설정

sSiteCode = ""			'IPIN 서비스 사이트 코드		(NICE신용평가정보에서 발급한 사이트코드)
sSitePw   = ""			'IPIN 서비스 사이트 패스워드	(NICE신용평가정보에서 발급한 사이트패스워드)


'사용자 정보 및 CP 요청번호를 암호화한 데이타입니다.
sResponseData = Fn_checkXss(Request("enc_data"), "encodeData")
Dim param_r1 : param_r1 = Fn_checkXss(Request("param_r1"), "")

'CP 요청번호 : ipin_main.asp 에서 세션 처리한 데이타
sCPRequest = SESSION("CPREQUEST")    

'IPIN 서비스 객체 정의
IPIN_DLL		= "IPINClient.Kisinfo"

'컴포넌트 객체 생성
SET clsIPINDll	= Server.CreateObject(IPIN_DLL)


'┌ 복호화 함수 설명  ──────────────────────────────────────────────────────────
'	Method 결과값(iRtn)에 따라, 프로세스 진행여부를 파악합니다.

'	fnResponse 함수는 결과 데이타를 복호화 하는 함수이며,
'	fnResponseExt 함수는 결과 데이타 복호화 및 CP 요청번호 일치여부도 확인하는 함수입니다. (세션에 넣은 sCPRequest 데이타로 검증)

'	따라서 귀사에서 원하는 함수로 이용하시기 바랍니다.
'└────────────────────────────────────────────────────────────────────
iRtn = clsIPINDll.fnResponse(sSiteCode, sSitePw, sResponseData)
'iRtn = clsIPINDll.fnResponseExt(sSiteCode, sSitePw, sResponseData, sCPRequest)

'Method 결과값에 따른 처리사항
IF (iRtn = 1) THEN

	'다음과 같이 사용자 정보를 추출할 수 있습니다.
	'사용자에게 보여주는 정보는, '이름' 데이타만 노출 가능합니다.
	
	'사용자 정보를 다른 페이지에서 이용하실 경우에는
	'보안을 위하여 암호화 데이타(sResponseData)를 통신하여 복호화 후 이용하실것을 권장합니다. (현재 페이지와 같은 처리방식)
	
	'만약, 복호화된 정보를 통신해야 하는 경우엔 데이타가 유출되지 않도록 주의해 주세요. (세션처리 권장)
	'form 태그의 hidden 처리는 데이타 유출 위험이 높으므로 권장하지 않습니다.
	
	sVNumber		= clsIPINDll.bstrVNumber				'가상주민번호 (13자리이며, 숫자 또는 문자 포함)
	sName			= clsIPINDll.bstrName					'이름
	sDupInfo		= clsIPINDll.bstrDupInfo				'중복가입 확인값 (DI - 64 byte 고유값)
	sAgeCode		= clsIPINDll.bstrAgeCode				'연령대 코드 (개발 가이드 참조)
	sGenderCode		= clsIPINDll.bstrGenderCode				'성별 코드 (개발 가이드 참조)
	sBirthDate		= clsIPINDll.bstrBirthDate				'생년월일 (YYYYMMDD)
	sNationalInfo	= clsIPINDll.bstrNationalInfo			'내/외국인 정보 (개발 가이드 참조)
	sCPRequestNum	= clsIPINDll.bstrCPRequestNUM			'CP 요청번호
	
	'sAuthInfo		= clsIPINDll.bstrAuthInfo				'본인확인 수단 (개발 가이드 참조)
	'sCoInfo1		= clsIPINDll.bstrCoInfo1				'연계정보 확인값 (CI - 88 byte 고유값)
	'sCIUpdate		= clsIPINDll.bstrCIUpdate				'CI 갱신정보
	
	'RESPONSE.WRITE "가상주민번호 : " & sVNumber & "<BR>"
	'RESPONSE.WRITE "이름 : " & sName & "<BR>"
	'RESPONSE.WRITE "중복가입 확인값 (DI) : " & sDupInfo & "<BR>"
	'RESPONSE.WRITE "연령대 코드 : " & sAgeCode & "<BR>"
	'RESPONSE.WRITE "성별 코드 : " & sGenderCode & "<BR>"
	'RESPONSE.WRITE "생년월일 : " & sBirthDate & "<BR>"
	'RESPONSE.WRITE "내/외국인 정보 : " & sNationalInfo & "<BR>"
	'RESPONSE.WRITE "CP 요청번호 : " & sCPRequestNum & "<BR>"
	'RESPONSE.WRITE "본인확인 수단 : " & sAuthInfo & "<BR>"
	'RESPONSE.WRITE "연계정보 확인값 (CI) : " & sCoInfo1 & "<BR>"
	'RESPONSE.WRITE "CI 갱신정보 : " & sCIUpdate & "<BR>"
	'RESPONSE.WRITE "------ 복호화 된 정보가 정상인지 확인해 주시기 바랍니다."
	'RESPONSE.WRITE "<BR><BR><BR><BR><BR><BR>"
	
	sRtnMsg = "정상 처리되었습니다."
	session("sVNumber")      = sVNumber
	session("sName")         = sName
	session("sBirthDate")    = sBirthDate
	session("sGender")       = sGenderCode
	session("sNationalInfo") = sNationalInfo
	session("sDupInfo")      = sDupInfo
	session("sConnInfo")     = sCoInfo1

ELSEIF (iRtn = -9) THEN
	sRtnMsg = "입력값 오류 : fnResponse 또는 fnResponseExt 함수 처리시, 필요한 파라미터값의 정보를 정확하게 입력해 주시기 바랍니다."
ELSEIF (iRtn = -12) THEN
	sRtnMsg = "CP 비밀번호 불일치 : IPIN 서비스 사이트 패스워드를 확인해 주시기 바랍니다."
ELSEIF (iRtn = -13) THEN
	sRtnMsg = "CP 요청번호 불일치 : 세션에 넣은 sCPRequest 데이타를 확인해 주시기 바랍니다."
ELSE
	sRtnMsg = "iRtn 값 확인 후, NICE신용평가정보 개발 담당자에게 문의해 주세요."
END IF

SET clsIPINDll = NOTHING

Function Fn_checkXss (CheckString, CheckGubun) 
	CheckString = trim(CheckString)
	CheckString = replace(CheckString,"<","&lt;")
	CheckString = replace(CheckString,">","&gt;")
	CheckString = replace(CheckString,"""","")  
	CheckString = replace(CheckString,"'","")   
	CheckString = replace(CheckString,"(","")
	CheckString = replace(CheckString,")","")
	CheckString = replace(CheckString,"#","")
	CheckString = replace(CheckString,"%","")
	CheckString = replace(CheckString,";","")
	CheckString = replace(CheckString,":","")
	CheckString = replace(CheckString,"-","")      
	CheckString = replace(CheckString,"`","")
	CheckString = replace(CheckString,"--","")
	CheckString = replace(CheckString,"\","")
	IF CheckGubun <> "encodeData" THEN	
		CheckString = replace(CheckString,"+","")
		CheckString = replace(CheckString,"=","")
		CheckString = replace(CheckString,"/","")
	END IF	
	Fn_checkXss = CheckString
End Function

If param_r1 <> "fPwd" Then 
	
	' 회원 중복검사
	Call Expires()
	Call dbopen()
		Call Check()
	Call dbclose()

	If FV_UserDelfg = "0" Then
		sRtnMsg = "이미 가입된 회원정보입니다."
		iRtn = -50
		GoUrl = "../member/login.asp"
	Else
		If FV_UserReJoin < "0" And NoJoinDate > "0" Then 
			sRtnMsg = "탈퇴 후 " & NoJoinDate & "일동안 회원가입이 불가능 하며,\n\n회원님께서는 " & Abs( int(FV_UserReJoin) ) & "일 후에 회원가입이 가능합니다.');"
			iRtn = -51
		End If
	End If
Else
	GoUrl = "fPwdResult.asp"
End If

Sub Check()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SELECT top 1 [UserDelfg] , ISNULL( datediff(d,dateadd(dd, ? ,[UserDelfgDate]),getdate() ) ,0) AS [UserReJoin]"  &_
	" FROM [dbo].[SP_USER_MEMBER] WHERE [UserDIKEY] = ? order by [UserIdx] desc"

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@NoJoinDate" ,adInteger , adParamInput, 0, NoJoinDate )
		.Parameters.Append .CreateParameter( "@DIKEY"      ,advarchar , adParamInput, 64, sDupInfo )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FV")
	Set objRs = Nothing
End Sub
		
%>

<html>
<head>
	<title>NICE신용평가정보 가상주민번호 서비스</title>
	<style type="text/css">
	BODY
	{
		COLOR: #7f7f7f;
		FONT-FAMILY: "Dotum","DotumChe","Arial";
		BACKGROUND-COLOR: #ffffff;
	}
	</style>
	<script language='javascript'>
	if('<%=iRtn%>' == '1'){
		alert('인증되었습니다.');
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "인증결과 : 정상적으로 인증이 완료 되었습니다. 다음단계를 진행해주세요! "
		parent.opener.parent.document.fm.authResult.value = "ipin";
		parent.opener.parent.document.fm.action = "<%=GoUrl%>";
		parent.opener.parent.document.fm.submit();
		self.close();
	}else if('<%=iRtn%>' == '-50'){
		alert('<%=sRtnMsg%>');
		parent.opener.parent.location.href = "<%=GoUrl%>";
		self.close();
	}else{
		alert('<%=sRtnMsg%>');
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "인증결과 : <%=sRtnMsg%>"
		self.close();
	}
	</script>
</head>

<body>
<!-- iRtn : <%= iRtn %> - <%= sRtnMsg %><br><br-->

<!-- 사용자 정보는 '이름' 외에는 화면에 노출하시면 안됩니다.
	 사용자 정보를 통신해야 하는 경우엔, 아래와 같이 암호화 정보로 통신 후 복호화하여 이용하시기 바랍니다.
	 만약, 복호화 된 데이타를 통신해야 하는 경우에는 정보보안을 위하여 주의해 주시기 바랍니다. -->
	 
<!--table border="0">
<tr>
	<td>이름 : <%= sName %></td>
</tr>

<form name="user" method="post">
	<input type="hidden" name="enc_data" value="<%= sResponseData %>"><br>
</form>
</table-->

</body>
</html>