<%
Dim clsCPClient
Dim sSiteCode, sSitePassword, sCipherTime
Dim sRequestNumber             '요청 번호
Dim sErrorCode                 '인증 결과코드
Dim sAuthType                  '인증 수단
Dim sReserved1, sReserved2, sReserved3
Dim sResult

Dim err_mgs

sEncodeData = Fn_checkXss(Request("EncodeData"), "encodeData")
sReserved1 = Fn_checkXss(Request("param_r1"), "")
sReserved2 = Fn_checkXss(Request("param_r2"), "")
sReserved3 = Fn_checkXss(Request("param_r3"), "")

sSiteCode      	= ""				'NICE로부터 부여받은 사이트 코드
sSitePassword   = ""			'NICE로부터 부여받은 사이트 패스워드

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
		If TRIM(sResult.Item("ERR_CODE")) = "0001" Then 
			err_mgs = "본인명의 휴대폰이 아니거나 잘못 입력하셨습니다.\n\n통신사와 입력정보를 확인해주시기 바랍니다."
		Else
			err_mgs = "본인인증 실패 사유 : " & TRIM(sResult.Item("ERR_CODE"))
		End If

		'RESPONSE.WRITE "인증결과_복호화_성공_원문[" & sPlain & "]<br>"
	
	  'RESPONSE.WRITE "요청 번호 : " & TRIM(sResult.Item("REQ_SEQ")) &"<br>"
	  'RESPONSE.WRITE "본인인증 실패 사유 : " & TRIM(sResult.Item("ERR_CODE")) &"<br>"
	  'RESPONSE.WRITE "인증수단 : " & TRIM(sResult.Item("AUTH_TYPE")) &"<br>"
		
	  sRequestNumber = TRIM(sResult.Item("REQ_SEQ"))
		sErrorCode = TRIM(sResult.Item("ERR_CODE"))
		sAuthType = TRIM(sResult.Item("AUTH_TYPE"))
	
	END IF	    
	
	sRequestNO = TRIM(sResult.Item("REQ_SEQ"))
	IF session("REQ_SEQ") <> sRequestNO THEN
		'RESPONSE.WRITE "세션값이 다릅니다. 올바른 경로로 접근하시기 바랍니다.<br>"
		err_mgs = "세션값이 다릅니다. 올바른 경로로 접근하시기 바랍니다."
	END IF
ELSE
	'RESPONSE.WRITE "요청정보_암호화_오류:" & iRtn & "<br>"
	' -1 : 암호화 시스템 에러입니다.
	' -4 : 입력 데이터 오류입니다.
	' -5 : 복호화 해쉬 오류입니다.
	' -6 : 복호화 데이터 오류입니다.
	' -9 : 입력 데이터 오류입니다.
	'-12 : 사이트 패스워드 오류입니다.
	err_mgs = "err code : " & iRtn & " 암호화 시스템 에러입니다. 잠시후에 시도해 주시기 바랍니다."
END IF

Set clsCPClient = Nothing
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
	
			If strName = "NAME" And bLen Mod 2 = 0 Then
				strFChar = Mid(str, nPos2+1, 1)
				
				Set ObjRegExp = New Regexp
				ObjRegExp.IgnoreCase = True
				ObjRegExp.Global = True
				ObjRegExp.Pattern = "[^-a-zA-Z0-9/]"
				Set MatchCnt = ObjRegExp.Execute(strFChar)
	
				If MatchCnt.count > 0 Then
					bLen = bLen / 2
				End If
	
				Set ObjRegExp = Nothing
				Set MatchCnt = Nothing
			End IF
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
	
END FUNCTION

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
	<script type="text/javascript">
	alert('<%=err_mgs%>');
	self.close();
	</script>
</head>
<body>
    <center>
    <p><p><p><p>
    본인인증이 실패하였습니다.<br>
    <table width=500 border=1>
        <tr>
            <td>복호화한 시간</td>
            <td><%= sCipherTime %> (YYMMDDHHMMSS)</td>
        </tr>
        <tr>
            <td>요청 번호</td>
            <td><%= sRequestNumber %></td>
        </tr>            
        <tr>
            <td>본인인증 실패 사유</td>
            <td><%= sErrorCode %></td>
        </tr>            
        <tr>
            <td>인증수단</td>
            <td><%= sAuthType %></td>
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
    </center>
</body>
</html>