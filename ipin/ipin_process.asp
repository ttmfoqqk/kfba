<%
	'********************************************************************************************************************************************
	'NICE�ſ������� Copyright(c) KOREA INFOMATION SERVICE INC. ALL RIGHTS RESERVED
	'
	'���񽺸� : �����ֹι�ȣ���� (IPIN) ����
	'�������� : �����ֹι�ȣ���� (IPIN) ����� ���� ���� ó�� ������
	
	'���Ź��� ������(�������)�� ����ȭ������ �ǵ����ְ�, close�� �ϴ� ��Ȱ�� �մϴ�.
  '********************************************************************************************************************************************
    
    DIM sResponseData, sReservedParam1, sReservedParam2, sReservedParam3
    
    '����� ���� �� CP ��û��ȣ�� ��ȣȭ�� ����Ÿ�Դϴ�. (ipin_main.asp ���������� ��ȣȭ�� ����Ÿ�ʹ� �ٸ��ϴ�.)
    sResponseData = Fn_checkXss(Request("enc_data"), "encodeData")
    
    'ipin_main.asp ���������� ������ ����Ÿ�� �ִٸ�, �Ʒ��� ���� Ȯ�ΰ����մϴ�.
    sReservedParam1 = Fn_checkXss(Request("param_r1"), "")
    sReservedParam2 = Fn_checkXss(Request("param_r2"), "")
    sReservedParam3 = Fn_checkXss(Request("param_r3"), "")
    
    '��ȣȭ�� ����� ������ �����ϴ� ���
    IF (sResponseData <> "") THEN
%>

<html>
<head>
	<title>NICE�ſ������� �����ֹι�ȣ ����</title>
	<script language='javascript'>
		function fnLoad()
		{
			// ��翡���� �ֻ����� �����ϱ� ���� 'parent.opener.parent.document.'�� �����Ͽ����ϴ�.
			// ���� �ͻ翡 ���μ����� �°� �����Ͻñ� �ٶ��ϴ�.
			parent.opener.parent.document.vnoform.enc_data.value = "<%= sResponseData %>";
			
			parent.opener.parent.document.vnoform.param_r1.value = "<%= sReservedParam1 %>";
			parent.opener.parent.document.vnoform.param_r2.value = "<%= sReservedParam2 %>";
			parent.opener.parent.document.vnoform.param_r3.value = "<%= sReservedParam3 %>";
			
			parent.opener.parent.document.vnoform.target = "Parent_window";
			
			// ���� �Ϸ�ÿ� ��������� �����ϰ� �Ǵ� �ͻ� Ŭ���̾�Ʈ ��� ������ URL
			parent.opener.parent.document.vnoform.action = "ipin_result.asp";
			parent.opener.parent.document.vnoform.submit();
			
			self.close();
		}
	</script>
</head>
<body onLoad="fnLoad()">

<%
	ELSE
%>

<html>
<head>
	<title>NICE�ſ������� �����ֹι�ȣ ����</title>
	<body onLoad="self.close()">

<%
	END IF
	
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

</body>
</html>