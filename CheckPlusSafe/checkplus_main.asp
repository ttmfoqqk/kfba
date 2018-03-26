<%
Dim clsCPClient
Dim iRtn, sEncData, sPlainData
Dim sRequestNO, sSiteCode, sSitePassword , sReturnUrl , sErrorUrl, popgubun, customize

Dim pageMode : pageMode = Request("pageMode")

SET clsCPClient  = SERVER.CREATEOBJECT("CPClient.NiceID")

sSiteCode       = ""			'NICE�κ��� �ο����� ����Ʈ �ڵ�
sSitePassword   = ""			'NICE�κ��� �ο����� ����Ʈ �н�����
sAuthType       = "M"		'������ �⺻ ����ȭ��, M: �ڵ���, C: ī��, X: ����������
popgubun 	    = "N"		'Y : ��ҹ�ư ���� / N : ��ҹ�ư ����
customize       = ""		'������ �⺻ �������� / Mobile : �����������
   
'CheckPlus(��������) ó�� ��, ��� ����Ÿ�� ���� �ޱ����� ���������� ���� http���� �Է��մϴ�.
sReturnUrl      = "http://" & Request.ServerVariables("SERVER_NAME") & "/CheckPlusSafe/checkplus_success.asp"			'������ �̵��� URL
sErrorUrl	    = "http://" & Request.ServerVariables("SERVER_NAME") & "/CheckPlusSafe/checkplus_fail.asp"				'���н� �̵��� URL

sRequestNO = "REQ0000000001"			'��û ��ȣ, �̴� ����/�����Ŀ� ���� ������ �ǵ����ְ� �ǹǷ�
								'��ü�� �����ϰ� �����Ͽ� ���ų�, �Ʒ��� ���� �����Ѵ�.
iRtn = clsCPClient.fnRequestNO(sSiteCode)

IF iRtn = 0 THEN
	sRequestNO = clsCPClient.bstrRandomRequestNO
	session("REQ_SEQ") = sRequestNO		'��ŷ���� ������ ���Ͽ� ������ ����Ѵٸ�, ���ǿ� ��û��ȣ�� �ִ´�.
END IF

sPlainData = fnGenPlainData(sRequestNO, sSiteCode, sAuthType, sReturnUrl, sErrorUrl, popgubun, customize)

'�������� ��ȣȭ
iRtn = clsCPClient.fnEncode(sSiteCode, sSitePassword, sPlainData)

IF iRtn = 0 THEN
	sEncData = clsCPClient.bstrCipherData
ELSE
	RESPONSE.WRITE "��û����_��ȣȭ_����:" & iRtn & "<br>"
	' -1 : ��ȣȭ �ý��� �����Դϴ�.
	' -2 : ��ȣȭ ó�������Դϴ�.
	' -3 : ��ȣȭ ������ �����Դϴ�.
	' -4 : �Է� ������ �����Դϴ�.
END IF

Set clsCPClient = Nothing
%>

<%
'**************************************************************************************
'���ڿ� ���� 
'**************************************************************************************  					          	
Function fnGenPlainData(aRequestNO, aSiteCode, aAuthType, aReturnUrl, aErrorUrl, popgubun, customize)
			
	'�Է� �Ķ���ͷ� plaindata ���� 			
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
'�Է��Ķ������ ���ڿ����� ����	
'**************************************************************************************  					          	
Function fnGetDataLength(aData)		
	Dim iData_len
	if (len(aData) > 0) then
		for i = 1 to len(aData)
			if (ASC(mid(aData,i,1)) < 0) then	'�ѱ��ΰ��
				iData_len = iData_len + 2
			else			'�ѱ��̾ƴѰ��
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
	<title>NICE�ſ������� - CheckPlus �Ƚɺ������� �׽�Ʈ</title>
	
	<script language='javascript'>
	window.onload = function(){
		fnPopup();
	};
	
	function fnPopup(){
		if('<%=iRtn%>' != '0'){alert('err code : <%=iRtn%> ��ȣȭ �ý��� �����Դϴ�. ����Ŀ� �õ��� �ֽñ� �ٶ��ϴ�.');self.close();return false;}
		document.form_chk.action = "https://nice.checkplus.co.kr/CheckPlusSafeModel/checkplus.cb";
		document.form_chk.submit();
	}
	</script>
</head>
<body>
	<!-- �������� ���� �˾��� ȣ���ϱ� ���ؼ��� ������ ���� form�� �ʿ��մϴ�. -->
	<form name="form_chk" method="post">
		<input type="hidden" name="m" value="checkplusSerivce">						<!-- �ʼ� ����Ÿ��, �����Ͻø� �ȵ˴ϴ�. -->
		<input type="hidden" name="EncodeData" value="<%= sEncData %>">		<!-- ������ ��ü������ ��ȣȭ �� ����Ÿ�Դϴ�. -->
	    
	    <!-- ��ü���� ����ޱ� ���ϴ� ����Ÿ�� �����ϱ� ���� ����� �� ������, ������� ����� �ش� ���� �״�� �۽��մϴ�.
	    	   �ش� �Ķ���ʹ� �߰��Ͻ� �� �����ϴ�. -->
		<input type="hidden" name="param_r1" value="<%=pageMode%>">
		<input type="hidden" name="param_r2" value="">
		<input type="hidden" name="param_r3" value="">
	</form>
</body>
</html>