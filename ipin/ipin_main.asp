<%
	'********************************************************************************************************************************************
	'NICE�ſ������� Copyright(c) KOREA INFOMATION SERVICE INC. ALL RIGHTS RESERVED
	'
	'���񽺸� : �����ֹι�ȣ���� (IPIN) ����
	'�������� : �����ֹι�ȣ���� (IPIN) ȣ�� ������
    '********************************************************************************************************************************************
    
    DIM sSiteCode, sSitePw, sReturnURL, sCPRequest
    DIM IPIN_DLL, clsIPINDll
    DIM iRtn, sEncReqData, sRtnMsg
	Dim pageMode : pageMode = Request("pageMode")

	sSiteCode = ""			'IPIN ���� ����Ʈ �ڵ�		(NICE�ſ����������� �߱��� ����Ʈ�ڵ�)
	sSitePw   = ""		'IPIN ���� ����Ʈ �н�����	(NICE�ſ����������� �߱��� ����Ʈ�н�����)
	
	
	
	'�� sReturnURL ������ ���� ����  ����������������������������������������������������������������������������������������������������������
	'	NICE�ſ������� �˾����� �������� ����� ������ ��ȣȭ�Ͽ� �ͻ�� �����մϴ�.
	'	���� ��ȣȭ�� ��� ����Ÿ�� ���Ϲ����� URL ������ �ּ���.
	
	'	* URL �� http ���� �Է��� �ּž��ϸ�, �ܺο����� ������ ��ȿ�� �������� �մϴ�.
	'	* ��翡�� �����ص帰 ���������� ��, ipin_process.asp �������� ����� ������ ���Ϲ޴� ���� �������Դϴ�.
	
	'	�Ʒ��� URL �����̸�, �ͻ��� ���� �����ΰ� ������ ���ε� �� ���������� ��ġ�� ���� ��θ� �����Ͻñ� �ٶ��ϴ�.
	'	�� - http://www.test.co.kr/ipin_process.asp, https://www.test.co.kr/ipin_process.asp, https://test.co.kr/ipin_process.asp
	'������������������������������������������������������������������������������������������������������������������������������������������
	sReturnURL				= "http://" & Request.ServerVariables("SERVER_NAME") & "/ipin/ipin_result.asp"
	
	
	'�� sCPRequest ������ ���� ����  ����������������������������������������������������������������������������������������������������������
	'	[CP ��û��ȣ]�� �ͻ翡�� ����Ÿ�� ���Ƿ� �����ϰų�, ��翡�� ������ ���� ����Ÿ�� ������ �� �ֽ��ϴ�.
	
	'	CP ��û��ȣ�� ���� �Ϸ� ��, ��ȣȭ�� ��� ����Ÿ�� �Բ� �����Ǹ�
	'	����Ÿ ������ ���� �� Ư�� ����ڰ� ��û�� ������ Ȯ���ϱ� ���� �������� �̿��Ͻ� �� �ֽ��ϴ�.
	
	'	���� �ͻ��� ���μ����� �����Ͽ� �̿��� �� �ִ� ����Ÿ�̱⿡, �ʼ����� �ƴմϴ�.
	'������������������������������������������������������������������������������������������������������������������������������������������
	sCPRequest				= ""
	
	
	
	
	
	'IPIN ���� ��ü ����
	IPIN_DLL		= "IPINClient.Kisinfo"
	
	'������Ʈ ��ü ����
	SET clsIPINDll	= Server.CreateObject(IPIN_DLL)
	
	
	'�ռ� ����帰 �ٿͰ���, CP ��û��ȣ�� ������ ����� ���� �Ʒ��� ���� ������ �� �ֽ��ϴ�.
	clsIPINDll.fnRequestSEQ(sSiteCode)
	sCPRequest = clsIPINDll.bstrRandomRequestSEQ
	
	'CP ��û��ȣ�� ���ǿ� �����մϴ�.
	'���� ������ ������ ������ ipin_result.asp ���������� ����Ÿ ������ ������ ���� Ȯ���ϱ� �����Դϴ�.
	'�ʼ������� �ƴϸ�, ������ ���� �ǰ�����Դϴ�.
	SESSION("CPREQUEST") = sCPRequest
	
	
	'Method �����(iRtn)�� ����, ���μ��� ���࿩�θ� �ľ��մϴ�.
	iRtn = clsIPINDll.fnRequest(sSiteCode, sSitePw, sCPRequest, sReturnURL)
	
	'Method ������� ���� ó������
	IF (iRtn = 0) THEN
	
		'fnRequest �Լ� ó���� ��ü������ ��ȣȭ�� �����͸� �����մϴ�.
		'����� ��ȣȭ�� ����Ÿ�� ��� �˾� ��û��, �Բ� �����ּž� �մϴ�.
		sEncReqData = clsIPINDll.bstrRequestCipherData
		
		sRtnMsg = "���� ó���Ǿ����ϴ�."
	
	ELSEIF (iRtn = -9) THEN
    	sRtnMsg = "�Է°� ���� : fnRequest �Լ� ó����, �ʿ��� 4���� �Ķ���Ͱ��� ������ ��Ȯ�ϰ� �Է��� �ֽñ� �ٶ��ϴ�."
    	sEncReqData = ""
    ELSE
    	sRtnMsg = "iRtn �� Ȯ�� ��, NICE�ſ������� ���� ����ڿ��� ������ �ּ���."
    	sEncReqData = ""
    END IF
    
    SET clsIPINDll = NOTHING

%>

<html>
<head>
	<title>NICE�ſ������� �����ֹι�ȣ ����</title>
	
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
<!-- �����ֹι�ȣ ���� �˾��� ȣ���ϱ� ���ؼ��� ������ ���� form�� �ʿ��մϴ�. -->
<form name="form_ipin" method="post">
	<input type="hidden" name="m" value="pubmain">						<!-- �ʼ� ����Ÿ��, �����Ͻø� �ȵ˴ϴ�. -->
    <input type="hidden" name="enc_data" value="<%= sEncReqData %>">	<!-- ������ ��ü������ ��ȣȭ �� ����Ÿ�Դϴ�. -->
    
    <!-- ��ü���� ����ޱ� ���ϴ� ����Ÿ�� �����ϱ� ���� ����� �� ������, ������� ����� �ش� ���� �״�� �۽��մϴ�.
    	 �ش� �Ķ���ʹ� �߰��Ͻ� �� �����ϴ�. -->
    <input type="hidden" name="param_r1" value="<%=pageMode%>">
    <input type="hidden" name="param_r2" value="">
    <input type="hidden" name="param_r3" value="">
</form>
</body>
</html>