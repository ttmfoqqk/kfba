<!-- #include file = "../_lib/header.asp" -->
<%
Dim clsCPClient
Dim sSiteCode, sSitePassword, sCipherTime
Dim sRequestNumber             '��û ��ȣ
Dim sResponseNumber            '���� ������ȣ
Dim sAuthType                  '���� ����
Dim sName                      '����
Dim sDupInfo                   '�ߺ����� Ȯ�ΰ� (DI_64 byte)
Dim sConnInfo                  '�������� Ȯ�ΰ� (CI_88 byte)
Dim sBirthDate                 '����
Dim sGender		               '����
Dim sNationalInfo              '��/�ܱ��� ���� (����� �Ŵ��� ����)
Dim sMobileNo,sMobileCo        '�������� �޴��� ����
Dim sReserved1, sReserved2, sReserved3
Dim sResult
Dim NoJoinDate  : NoJoinDate    = 0 ' Ż���� �簡�� ������ ��¥ ����
Dim GoUrl       : GoUrl = "joinData.asp"

Dim err_mgs

sEncodeData = Fn_checkXss(Request("EncodeData"), "encodeData")
sReserved1 = Fn_checkXss(Request("param_r1"), "")
sReserved2 = Fn_checkXss(Request("param_r2"), "")
sReserved3 = Fn_checkXss(Request("param_r3"), "")

	sSiteCode     	= "G5917"			'NICE�κ��� �ο����� ����Ʈ �ڵ�
	sSitePassword   = "KDRKENUJI90J"			'NICE�κ��� �ο����� ����Ʈ �н�����

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
		
		'RESPONSE.WRITE "�������_��ȣȭ_����_����[" & sPlain & "]<br>"
	
		'RESPONSE.WRITE "��û ��ȣ : " & TRIM(sResult.Item("REQ_SEQ")) &"<br>"
		'RESPONSE.WRITE "���� ��ȣ : " & TRIM(sResult.Item("RES_SEQ")) &"<br>"
		'RESPONSE.WRITE "�������� : " & TRIM(sResult.Item("AUTH_TYPE")) &"<br>"
		'RESPONSE.WRITE "���� : " & TRIM(sResult.Item("NAME")) &"<br>"
		'RESPONSE.WRITE "���� : " & TRIM(sResult.Item("BIRTHDATE")) &"<br>"
		'RESPONSE.WRITE "���� : " & TRIM(sResult.Item("GENDER")) &"<br>"
		'RESPONSE.WRITE "��/�ܱ��� ���� : " & TRIM(sResult.Item("NATIONALINFO")) &"<br>"
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
		'RESPONSE.WRITE "���ǰ��� �ٸ��ϴ�. �ùٸ� ��η� �����Ͻñ� �ٶ��ϴ�.<br>"
		err_mgs = "���ǰ��� �ٸ��ϴ�. �ùٸ� ��η� �����Ͻñ� �ٶ��ϴ�."
	END IF
		
	Else
	err_mgs = "err code : " & iRtn & " ��ȣȭ �ý��� �����Դϴ�. ����Ŀ� �õ��� �ֽñ� �ٶ��ϴ�."
	'RESPONSE.WRITE "��û����_��ȣȭ_����:" & iRtn & "<br>"
	' -1 : ��ȣȭ �ý��� �����Դϴ�.
	' -4 : �Է� ������ �����Դϴ�.
	' -5 : ��ȣȭ �ؽ� �����Դϴ�.
	' -6 : ��ȣȭ ������ �����Դϴ�.
	' -9 : �Է� ������ �����Դϴ�.
  '-12 : ����Ʈ �н����� �����Դϴ�.
END IF

Set clsCPClient = Nothing


If sReserved1 <> "fPwd" Then 
	' ȸ�� �ߺ��˻�
	Call Expires()
	Call dbopen()
		Call Check()
	Call dbclose()

	If FV_UserDelfg = "0" Then
		err_mgs = "�̹� ���Ե� ȸ�������Դϴ�."
		iRtn = -50
		GoUrl = "../member/login.asp"
	Else
		If FV_UserReJoin < "0" And NoJoinDate > "0" Then 
			err_mgs = "Ż�� �� " & NoJoinDate & "�ϵ��� ȸ�������� �Ұ��� �ϸ�,\n\nȸ���Բ����� " & Abs( int(FV_UserReJoin) ) & "�� �Ŀ� ȸ�������� �����մϴ�.');"
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
    <title>NICE�ſ������� - CheckPlus �Ƚɺ�������</title>
	<script language='javascript'>
	if('<%=iRtn%>' == '0'){
		alert('�����Ǿ����ϴ�.');
		parent.opener.parent.document.fm.authResult.value = "safe";
		parent.opener.parent.document.fm.action = "<%=GoUrl%>";
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "������� : ���������� ������ �Ϸ� �Ǿ����ϴ�. �����ܰ踦 �������ּ���! "
		parent.opener.parent.document.fm.submit();
		self.close();
	}else if('<%=iRtn%>' == '-50'){
		alert('<%=err_mgs%>');
		parent.opener.parent.location.href = "<%=GoUrl%>";
		self.close();
	}else{
		alert('<%=err_mgs%>');
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "������� : <%=err_mgs%>"
		self.close();
	}
	</script>
</head>
<body>
    <!--center>
    <p><p><p><p>
    ���������� �Ϸ� �Ǿ����ϴ�.<br>
    <table border=1>
        <tr>
            <td>��ȣȭ�� �ð�</td>
            <td><%= sCipherTime %> (YYMMDDHHMMSS)</td>
        </tr>
        <tr>
            <td>��û ��ȣ</td>
            <td><%= sRequestNumber %></td>
        </tr>            
        <tr>
            <td>NICE���� ��ȣ</td>
            <td><%= sResponseNumber %></td>
        </tr>            
        <tr>
            <td>��������</td>
            <td><%= sAuthType %></td>
        </tr>
        <tr>
            <td>����</td>
            <td><%= sName %></td>
        </tr>
				<tr>
            <td>����</td>
            <td><%= sGender %></td>
        </tr>
					<tr>
            <td>�������</td>
            <td><%= sBirthDate %></td>
        </tr>
				<tr>
            <td>��/�ܱ��� ����</td>
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