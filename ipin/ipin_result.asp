<!-- #include file = "../_lib/header.asp" -->
<%
'********************************************************************************************************************************************
'NICE�ſ������� Copyright(c) KOREA INFOMATION SERVICE INC. ALL RIGHTS RESERVED
'
'���񽺸� : �����ֹι�ȣ���� (IPIN) ����
'�������� : �����ֹι�ȣ���� (IPIN) ��� ������
'********************************************************************************************************************************************

DIM sSiteCode, sSitePw, sResponseData, sCPRequest
DIM IPIN_DLL, clsIPINDll
DIM iRtn, sRtnMsg
DIM sVNumber, sName, sDupInfo, sAgeCode, sGenderCode, sBirthDate, sNationalInfo, sCPRequestNum
Dim sCoInfo1, sCIUpdate ,sAuthInfo
Dim GoUrl : GoUrl = "joinData.asp"

Dim NoJoinDate  : NoJoinDate    = 0 ' Ż���� �簡�� ������ ��¥ ����

sSiteCode = ""			'IPIN ���� ����Ʈ �ڵ�		(NICE�ſ����������� �߱��� ����Ʈ�ڵ�)
sSitePw   = ""			'IPIN ���� ����Ʈ �н�����	(NICE�ſ����������� �߱��� ����Ʈ�н�����)


'����� ���� �� CP ��û��ȣ�� ��ȣȭ�� ����Ÿ�Դϴ�.
sResponseData = Fn_checkXss(Request("enc_data"), "encodeData")
Dim param_r1 : param_r1 = Fn_checkXss(Request("param_r1"), "")

'CP ��û��ȣ : ipin_main.asp ���� ���� ó���� ����Ÿ
sCPRequest = SESSION("CPREQUEST")    

'IPIN ���� ��ü ����
IPIN_DLL		= "IPINClient.Kisinfo"

'������Ʈ ��ü ����
SET clsIPINDll	= Server.CreateObject(IPIN_DLL)


'�� ��ȣȭ �Լ� ����  ��������������������������������������������������������������������������������������������������������������������
'	Method �����(iRtn)�� ����, ���μ��� ���࿩�θ� �ľ��մϴ�.

'	fnResponse �Լ��� ��� ����Ÿ�� ��ȣȭ �ϴ� �Լ��̸�,
'	fnResponseExt �Լ��� ��� ����Ÿ ��ȣȭ �� CP ��û��ȣ ��ġ���ε� Ȯ���ϴ� �Լ��Դϴ�. (���ǿ� ���� sCPRequest ����Ÿ�� ����)

'	���� �ͻ翡�� ���ϴ� �Լ��� �̿��Ͻñ� �ٶ��ϴ�.
'������������������������������������������������������������������������������������������������������������������������������������������
iRtn = clsIPINDll.fnResponse(sSiteCode, sSitePw, sResponseData)
'iRtn = clsIPINDll.fnResponseExt(sSiteCode, sSitePw, sResponseData, sCPRequest)

'Method ������� ���� ó������
IF (iRtn = 1) THEN

	'������ ���� ����� ������ ������ �� �ֽ��ϴ�.
	'����ڿ��� �����ִ� ������, '�̸�' ����Ÿ�� ���� �����մϴ�.
	
	'����� ������ �ٸ� ���������� �̿��Ͻ� ��쿡��
	'������ ���Ͽ� ��ȣȭ ����Ÿ(sResponseData)�� ����Ͽ� ��ȣȭ �� �̿��Ͻǰ��� �����մϴ�. (���� �������� ���� ó�����)
	
	'����, ��ȣȭ�� ������ ����ؾ� �ϴ� ��쿣 ����Ÿ�� ������� �ʵ��� ������ �ּ���. (����ó�� ����)
	'form �±��� hidden ó���� ����Ÿ ���� ������ �����Ƿ� �������� �ʽ��ϴ�.
	
	sVNumber		= clsIPINDll.bstrVNumber				'�����ֹι�ȣ (13�ڸ��̸�, ���� �Ǵ� ���� ����)
	sName			= clsIPINDll.bstrName					'�̸�
	sDupInfo		= clsIPINDll.bstrDupInfo				'�ߺ����� Ȯ�ΰ� (DI - 64 byte ������)
	sAgeCode		= clsIPINDll.bstrAgeCode				'���ɴ� �ڵ� (���� ���̵� ����)
	sGenderCode		= clsIPINDll.bstrGenderCode				'���� �ڵ� (���� ���̵� ����)
	sBirthDate		= clsIPINDll.bstrBirthDate				'������� (YYYYMMDD)
	sNationalInfo	= clsIPINDll.bstrNationalInfo			'��/�ܱ��� ���� (���� ���̵� ����)
	sCPRequestNum	= clsIPINDll.bstrCPRequestNUM			'CP ��û��ȣ
	
	'sAuthInfo		= clsIPINDll.bstrAuthInfo				'����Ȯ�� ���� (���� ���̵� ����)
	'sCoInfo1		= clsIPINDll.bstrCoInfo1				'�������� Ȯ�ΰ� (CI - 88 byte ������)
	'sCIUpdate		= clsIPINDll.bstrCIUpdate				'CI ��������
	
	'RESPONSE.WRITE "�����ֹι�ȣ : " & sVNumber & "<BR>"
	'RESPONSE.WRITE "�̸� : " & sName & "<BR>"
	'RESPONSE.WRITE "�ߺ����� Ȯ�ΰ� (DI) : " & sDupInfo & "<BR>"
	'RESPONSE.WRITE "���ɴ� �ڵ� : " & sAgeCode & "<BR>"
	'RESPONSE.WRITE "���� �ڵ� : " & sGenderCode & "<BR>"
	'RESPONSE.WRITE "������� : " & sBirthDate & "<BR>"
	'RESPONSE.WRITE "��/�ܱ��� ���� : " & sNationalInfo & "<BR>"
	'RESPONSE.WRITE "CP ��û��ȣ : " & sCPRequestNum & "<BR>"
	'RESPONSE.WRITE "����Ȯ�� ���� : " & sAuthInfo & "<BR>"
	'RESPONSE.WRITE "�������� Ȯ�ΰ� (CI) : " & sCoInfo1 & "<BR>"
	'RESPONSE.WRITE "CI �������� : " & sCIUpdate & "<BR>"
	'RESPONSE.WRITE "------ ��ȣȭ �� ������ �������� Ȯ���� �ֽñ� �ٶ��ϴ�."
	'RESPONSE.WRITE "<BR><BR><BR><BR><BR><BR>"
	
	sRtnMsg = "���� ó���Ǿ����ϴ�."
	session("sVNumber")      = sVNumber
	session("sName")         = sName
	session("sBirthDate")    = sBirthDate
	session("sGender")       = sGenderCode
	session("sNationalInfo") = sNationalInfo
	session("sDupInfo")      = sDupInfo
	session("sConnInfo")     = sCoInfo1

ELSEIF (iRtn = -9) THEN
	sRtnMsg = "�Է°� ���� : fnResponse �Ǵ� fnResponseExt �Լ� ó����, �ʿ��� �Ķ���Ͱ��� ������ ��Ȯ�ϰ� �Է��� �ֽñ� �ٶ��ϴ�."
ELSEIF (iRtn = -12) THEN
	sRtnMsg = "CP ��й�ȣ ����ġ : IPIN ���� ����Ʈ �н����带 Ȯ���� �ֽñ� �ٶ��ϴ�."
ELSEIF (iRtn = -13) THEN
	sRtnMsg = "CP ��û��ȣ ����ġ : ���ǿ� ���� sCPRequest ����Ÿ�� Ȯ���� �ֽñ� �ٶ��ϴ�."
ELSE
	sRtnMsg = "iRtn �� Ȯ�� ��, NICE�ſ������� ���� ����ڿ��� ������ �ּ���."
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
	
	' ȸ�� �ߺ��˻�
	Call Expires()
	Call dbopen()
		Call Check()
	Call dbclose()

	If FV_UserDelfg = "0" Then
		sRtnMsg = "�̹� ���Ե� ȸ�������Դϴ�."
		iRtn = -50
		GoUrl = "../member/login.asp"
	Else
		If FV_UserReJoin < "0" And NoJoinDate > "0" Then 
			sRtnMsg = "Ż�� �� " & NoJoinDate & "�ϵ��� ȸ�������� �Ұ��� �ϸ�,\n\nȸ���Բ����� " & Abs( int(FV_UserReJoin) ) & "�� �Ŀ� ȸ�������� �����մϴ�.');"
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
	<title>NICE�ſ������� �����ֹι�ȣ ����</title>
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
		alert('�����Ǿ����ϴ�.');
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "������� : ���������� ������ �Ϸ� �Ǿ����ϴ�. �����ܰ踦 �������ּ���! "
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
		parent.opener.parent.document.getElementById('joinAuthTxtBox').innerHTML = "������� : <%=sRtnMsg%>"
		self.close();
	}
	</script>
</head>

<body>
<!-- iRtn : <%= iRtn %> - <%= sRtnMsg %><br><br-->

<!-- ����� ������ '�̸�' �ܿ��� ȭ�鿡 �����Ͻø� �ȵ˴ϴ�.
	 ����� ������ ����ؾ� �ϴ� ��쿣, �Ʒ��� ���� ��ȣȭ ������ ��� �� ��ȣȭ�Ͽ� �̿��Ͻñ� �ٶ��ϴ�.
	 ����, ��ȣȭ �� ����Ÿ�� ����ؾ� �ϴ� ��쿡�� ���������� ���Ͽ� ������ �ֽñ� �ٶ��ϴ�. -->
	 
<!--table border="0">
<tr>
	<td>�̸� : <%= sName %></td>
</tr>

<form name="user" method="post">
	<input type="hidden" name="enc_data" value="<%= sResponseData %>"><br>
</form>
</table-->

</body>
</html>