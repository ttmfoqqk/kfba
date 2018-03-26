<!-- #include file = "../_lib/header.asp" -->
<%
Dim clsCPClient
Dim sSiteCode, sSitePassword, sCipherTime
Dim sRequestNumber             '요청 번호
Dim sResponseNumber            '인증 고유번호
Dim sAuthType                  '인증 수단
Dim sName                      '성명
Dim sDupInfo                   '중복가입 확인값 (DI_64 byte)
Dim sConnInfo                  '연계정보 확인값 (CI_88 byte)
Dim sBirthDate                 '생일
Dim sGender		               '성별
Dim sNationalInfo              '내/외국인 정보 (사용자 매뉴얼 참조)
Dim sMobileNo,sMobileCo        '인증받은 휴대폰 정보
Dim sReserved1, sReserved2, sReserved3
Dim sResult
Dim NoJoinDate  : NoJoinDate    = 0 ' 탈퇴후 재가입 가능한 날짜 설정
Dim GoUrl       : GoUrl = "joinData.asp"

Dim err_mgs

sEncodeData = Fn_checkXss(Request("EncodeData"), "encodeData")
sReserved1 = Fn_checkXss(Request("param_r1"), "")
sReserved2 = Fn_checkXss(Request("param_r2"), "")
sReserved3 = Fn_checkXss(Request("param_r3"), "")

	sSiteCode     	= "G5917"			'NICE로부터 부여받은 사이트 코드
	sSitePassword   = "KDRKENUJI90J"			'NICE로부터 부여받은 사이트 패스워드

SET clsCPClient  = SERVER.CREATEOBJECT("CPClient.NiceID")
iRtn = clsCPClient.fnDecode(sSiteCode, sSitePassword, sEncodeData)

IF iRtn = 0 THEN
	sPlain           = clsCPClient.bstrPlainData
	sCipherTime      = clsCPClient.bstrCipherDateTime
	
	SET sResult = getParamData(sPlain)

	IF sResult IS NOTHING THEN     
		'RESPONSE.WRITE "응답값이 유효하지 않습니다."
		err_mgs = "응답값이 유효하지 않습니다."
		
	ELSE
		
		'RESPONSE.WRITE "인증결과_복호화_성공_원문[" & sPlain & "]<br>"
	
		'RESPONSE.WRITE "요청 번호 : " & TRIM(sResult.Item("REQ_SEQ")) &"<br>"
		'RESPONSE.WRITE "응답 번호 : " & TRIM(sResult.Item("RES_SEQ")) &"<br>"
		'RESPONSE.WRITE "인증수단 : " & TRIM(sResult.Item("AUTH_TYPE")) &"<br>"
		'RESPONSE.WRITE "성명 : " & TRIM(sResult.Item("NAME")) &"<br>"
		'RESPONSE.WRITE "생일 : " & TRIM(sResult.Item("BIRTHDATE")) &"<br>"
		'RESPONSE.WRITE "성별 : " & TRIM(sResult.Item("GENDER")) &"<br>"
		'RESPONSE.WRITE "내/외국인 정보 : " & TRIM(sResult.Item("NATIONALINFO")) &"<br>"
		'RESPONSE.WRITE "DI(64byte) : " & TRIM(sResult.Item("DI")) &"<br>"
		'RESPONSE.WRITE "CI(88byte) : " & TRIM(sResult.Item("CI")) &"<br>"

		sRequestNumber   = TRIM(sResult.Item("REQ_SEQ"))
		sResponseNumber  = TRIM(sResult.Item("RES_SEQ"))
		sAuthType        = TRIM(sResult.Item("AUTH_TYPE"))
		sName            = TRIM(sResult.Item("NAME"))
		sBirthDate       = TRIM(sResult.Item("BIRTHDATE"))
		sGender		     = TRIM(sResult.Item("GENDER"))
		sNationalInfo    = TRIM(sResult.Item("NATIONALINFO"))
		sDupInfo         = TRIM(sResult.Item("DI"))
		sConnInfo        = TRIM(sResult.Item("CI"))
		sMobileNo        = TRIM(sResult.Item("MOBILE_NO"))
		sMobileCo        = TRIM(sResult.Item("MOBILE_CO"))

		session("sName")         = sName
		session("sBirthDate")    = sBirthDate
		session("sGender")       = sGender
		session("sNationalInfo") = sNationalInfo
		session("sDupInfo")      = sDupInfo
		session("sConnInfo")     = sConnInfo
		session("sMobileNo")     = sMobileNo
		session("sMobileCo")     = sMobileCo
		
	END IF
	
	sRequestNO = TRIM(sResult.Item("REQ_SEQ"))
	
	IF session("REQ_SEQ") <> sRequestNO THEN
		'RESPONSE.WRITE "세션값이 다릅니다. 올바른 경로로 접근하시기 바랍니다.<br>"
		err_mgs = "세션값이 다릅니다. 올바른 경로로 접근하시기 바랍니다."
	END IF
		
	Else
	err_mgs = "err code : " & iRtn & " 암호화 시스템 에러입니다. 잠시후에 시도해 주시기 바랍니다."
	'RESPONSE.WRITE "요청정보_암호화_오류:" & iRtn & "<br>"
	' -1 : 암호화 시스템 에러입니다.
	' -4 : 입력 데이터 오류입니다.
	' -5 : 복호화 해쉬 오류입니다.
	' -6 : 복호화 데이터 오류입니다.
	' -9 : 입력 데이터 오류입니다.
  '-12 : 사이트 패스워드 오류입니다.
END IF

Set clsCPClient = Nothing


If sReserved1 <> "fPwd" Then 
	' 회원 중복검사
	Call Expires()
	Call dbopen()
		Call Check()
	Call dbclose()

	If FV_UserDelfg = "0" Then
		err_mgs = "이미 가입된 회원정보입니다."
		iRtn = -50
		GoUrl = "../member/login.asp"
	Else
		If FV_UserReJoin < "0" And NoJoinDate > "0" Then 
			err_mgs = "탈퇴 후 " & NoJoinDate & "일동안 회원가입이 불가능 하며,\n\n회원님께서는 " & Abs( int(FV_UserReJoin) ) & "일 후에 회원가입이 가능합니다.');"
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

<%
FUNCTION getParamData(pOrgData)

	DIM pDicData, arrData, iLoop

	arrData = f_getParamArray(pOrgData)
	IF ISARRAY(arrData) THEN 
		SET pDicData = SERVER.CREATEOBJECT("Scripting.Dictionary")
		IF UBOUND(arrData, 2) > 0 THEN
			FOR iLoop=1 TO UBOUND(arrData, 2)            
				pDicData.ADD arrData(1, iLoop), arrData(2, iLoop)
			NEXT
		END IF
		SET getParamData = pDicData
	ELSE 
		SET getParamData = NOTHING    
	END IF
END FUNCTION


FUNCTION f_getParamArray(str)

		DIM paramIndex, bLen, nPos1, nPos2
		DIM strName
		DIM strValue
		Dim strFChar, objRegExp, MatchCnt
		DIM paramList()
	
		paramIndex=1
		nPos1 = 1	' length의 시작 위치
		nPos2 = 1	' ":"의 위치
	
		DO WHILE nPos1 <= LEN(str)
		
			REDIM PRESERVE paramList(2, paramIndex)
	
			nPos2 = INSTR(nPos1, str, ":")
			bLen = MID(str, nPos1, (nPos2 - nPos1) )
			strName = f_getString(str, nPos2+1, bLen)	
			
			nPos1 = nPos2 + LEN(strName) + 1
			nPos2 = INSTR(nPos1, str, ":")
			bLen = MID(str, nPos1, (nPos2 - nPos1) )
	
			IF strName = "NAME" And bLen Mod 2 = 0 Then
				strFChar = Mid(str, nPos2+1, 1)
				
				Set ObjRegExp = New Regexp
				ObjRegExp.IgnoreCase = True
				ObjRegExp.Global = True
				ObjRegExp.Pattern = "[^-a-zA-Z0-9/]"
				Set MatchCnt = ObjRegExp.Execute(strFChar)
	
				IF MatchCnt.count > 0 THEN
					bLen = bLen / 2
				END IF
	
				Set ObjRegExp = Nothing
				Set MatchCnt = Nothing
			END IF
			strValue = f_getString(str, nPos2+1, bLen)	
			
			nPos1 = nPos2 + LEN(strValue) + 1
			
			paramList(1,paramIndex) = strName
			paramList(2,paramIndex) = strValue
			paramIndex = paramIndex +1
			
		LOOP
		
		IF paramIndex > 1 THEN 
			f_getParamArray = paramList
		ELSE
			f_getParamArray = ""
		END IF
		
	END FUNCTION


FUNCTION f_getString(s, start , bytesLen)

	Dim i, addLen

	addLen = 0
	for i=start to (start + bytesLen - 1) 
	 f_getString = f_getString & mid(s, i, 1)
	 addLen = addLen + 1
	 if Asc(mid(s, i, 1)) < 0 then
	  addLen = addLen + 1
	 end if
	 
	 if cint(addLen) = cint(bytesLen) then
	  exit for
	 end if  
	next 
	
end FUNCTION

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
%>

<html>
<head>
    <title>NICE신용평가정보 - CheckPlus 안심본인인증</title>
	<script language='javascript'>
	if('<%=iRtn%>' == '0'){
		alert('인증되었습니다.');
		parent.opener.parent.document.fm.authResult.value = "safe";
		parent.opener.parent.document.fm.action = "<%=GoUrl%>";
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "인증결과 : 정상적으로 인증이 완료 되었습니다. 다음단계를 진행해주세요! "
		parent.opener.parent.document.fm.submit();
		self.close();
	}else if('<%=iRtn%>' == '-50'){
		alert('<%=err_mgs%>');
		parent.opener.parent.location.href = "<%=GoUrl%>";
		self.close();
	}else{
		alert('<%=err_mgs%>');
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "인증결과 : <%=err_mgs%>"
		self.close();
	}
	</script>
</head>
<body>
    <!--center>
    <p><p><p><p>
    본인인증이 완료 되었습니다.<br>
    <table border=1>
        <tr>
            <td>복호화한 시간</td>
            <td><%= sCipherTime %> (YYMMDDHHMMSS)</td>
        </tr>
        <tr>
            <td>요청 번호</td>
            <td><%= sRequestNumber %></td>
        </tr>            
        <tr>
            <td>NICE응답 번호</td>
            <td><%= sResponseNumber %></td>
        </tr>            
        <tr>
            <td>인증수단</td>
            <td><%= sAuthType %></td>
        </tr>
        <tr>
            <td>성명</td>
            <td><%= sName %></td>
        </tr>
				<tr>
            <td>성별</td>
            <td><%= sGender %></td>
        </tr>
					<tr>
            <td>생년월일</td>
            <td><%= sBirthDate %></td>
        </tr>
				<tr>
            <td>내/외국인 정보</td>
            <td><%= sNationalInfo %></td>
        </tr>
					<tr>
            <td>DI</td>
            <td><%= sDupInfo %></td>
        </tr>
					<tr>
            <td>CI</td>
            <td><%= sConnInfo %></td>
        </tr>
        <tr>
            <td>RESERVED1</td>
            <td><%= sReserved1 %></td>
        </tr>
        <tr>
            <td>RESERVED2</td>
            <td><%= sReserved2 %></td>
        </tr>
        <tr>
            <td>RESERVED3</td>
            <td><%= sReserved3 %></td>
        </tr>
    </table>
    </center-->
</body>
</html>