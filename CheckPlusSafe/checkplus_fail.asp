<%
Dim clsCPClient
Dim sSiteCode, sSitePassword, sCipherTime
Dim sRequestNumber             '��û ��ȣ
Dim sErrorCode                 '���� ����ڵ�
Dim sAuthType                  '���� ����
Dim sReserved1, sReserved2, sReserved3
Dim sResult

Dim err_mgs

sEncodeData = Fn_checkXss(Request("EncodeData"), "encodeData")
sReserved1 = Fn_checkXss(Request("param_r1"), "")
sReserved2 = Fn_checkXss(Request("param_r2"), "")
sReserved3 = Fn_checkXss(Request("param_r3"), "")

sSiteCode      	= ""				'NICE�κ��� �ο����� ����Ʈ �ڵ�
sSitePassword   = ""			'NICE�κ��� �ο����� ����Ʈ �н�����

SET clsCPClient  = SERVER.CREATEOBJECT("CPClient.NiceID")
iRtn = clsCPClient.fnDecode(sSiteCode, sSitePassword, sEncodeData)

IF iRtn = 0 THEN
	sPlain           = clsCPClient.bstrPlainData
	sCipherTime      = clsCPClient.bstrCipherDateTime

	SET sResult = getParamData(sPlain)

	IF sResult IS NOTHING THEN     
		'RESPONSE.WRITE "���䰪�� ��ȿ���� �ʽ��ϴ�."
		err_mgs = "���䰪�� ��ȿ���� �ʽ��ϴ�."
	ELSE 
		If TRIM(sResult.Item("ERR_CODE")) = "0001" Then 
			err_mgs = "���θ��� �޴����� �ƴϰų� �߸� �Է��ϼ̽��ϴ�.\n\n��Ż�� �Է������� Ȯ�����ֽñ� �ٶ��ϴ�."
		Else
			err_mgs = "�������� ���� ���� : " & TRIM(sResult.Item("ERR_CODE"))
		End If

		'RESPONSE.WRITE "�������_��ȣȭ_����_����[" & sPlain & "]<br>"
	
	  'RESPONSE.WRITE "��û ��ȣ : " & TRIM(sResult.Item("REQ_SEQ")) &"<br>"
	  'RESPONSE.WRITE "�������� ���� ���� : " & TRIM(sResult.Item("ERR_CODE")) &"<br>"
	  'RESPONSE.WRITE "�������� : " & TRIM(sResult.Item("AUTH_TYPE")) &"<br>"
		
	  sRequestNumber = TRIM(sResult.Item("REQ_SEQ"))
		sErrorCode = TRIM(sResult.Item("ERR_CODE"))
		sAuthType = TRIM(sResult.Item("AUTH_TYPE"))
	
	END IF	    
	
	sRequestNO = TRIM(sResult.Item("REQ_SEQ"))
	IF session("REQ_SEQ") <> sRequestNO THEN
		'RESPONSE.WRITE "���ǰ��� �ٸ��ϴ�. �ùٸ� ��η� �����Ͻñ� �ٶ��ϴ�.<br>"
		err_mgs = "���ǰ��� �ٸ��ϴ�. �ùٸ� ��η� �����Ͻñ� �ٶ��ϴ�."
	END IF
ELSE
	'RESPONSE.WRITE "��û����_��ȣȭ_����:" & iRtn & "<br>"
	' -1 : ��ȣȭ �ý��� �����Դϴ�.
	' -4 : �Է� ������ �����Դϴ�.
	' -5 : ��ȣȭ �ؽ� �����Դϴ�.
	' -6 : ��ȣȭ ������ �����Դϴ�.
	' -9 : �Է� ������ �����Դϴ�.
	'-12 : ����Ʈ �н����� �����Դϴ�.
	err_mgs = "err code : " & iRtn & " ��ȣȭ �ý��� �����Դϴ�. ����Ŀ� �õ��� �ֽñ� �ٶ��ϴ�."
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
		nPos1 = 1	' length�� ���� ��ġ
		nPos2 = 1	' ":"�� ��ġ
	
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
    <title>NICE�ſ������� - CheckPlus �Ƚɺ�������</title>
	<script type="text/javascript">
	alert('<%=err_mgs%>');
	self.close();
	</script>
</head>
<body>
    <center>
    <p><p><p><p>
    ���������� �����Ͽ����ϴ�.<br>
    <table width=500 border=1>
        <tr>
            <td>��ȣȭ�� �ð�</td>
            <td><%= sCipherTime %> (YYMMDDHHMMSS)</td>
        </tr>
        <tr>
            <td>��û ��ȣ</td>
            <td><%= sRequestNumber %></td>
        </tr>            
        <tr>
            <td>�������� ���� ����</td>
            <td><%= sErrorCode %></td>
        </tr>            
        <tr>
            <td>��������</td>
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