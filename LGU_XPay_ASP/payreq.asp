<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<!-- #include file="./lgdacom/md5.asp" -->
<%
if session("UserIdx") = "" or IsNull(session("UserIdx")) Then
	With Response
	 .Write "<script type='text/javascript'>alert('�α����� �ʿ��մϴ�.');window.opener.location.reload();window.close()</script>"
	 .End
	End With
end If


Dim savePath : savePath = "./appMember/" '÷�� ������.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath
UPLOAD__FORM.MaxFileLen		= 20 * 1024 * 1024 '10�ް�






'With Response
' .Write "<script type='text/javascript'>alert('���� �ǰ����� �ƴ� ���� �׽�Ʈ�� �Դϴ�.\n\nȸ���Ե� ������ \'�ӽ� ��������\' �� �̵��ϼż� ���� ��Ź �帳�ϴ�.\n���ǻ����� 1800-6288������ �����ֽñ� �ٶ��ϴ�.\n\n������ ��� ����� �˼��մϴ�.');</script>"
'End With

Dim programIdx : programIdx = UPLOAD__FORM("programIdx")
Dim areaIdx    : areaIdx    = UPLOAD__FORM("areaIdx")
Dim payMethod  : payMethod  = UPLOAD__FORM("payMethod")

Dim LastName       : LastName       = UPLOAD__FORM("LastName")
Dim FirstName      : FirstName      = UPLOAD__FORM("FirstName")

Dim PhotoName      : PhotoName      = UPLOAD__FORM("PhotoName")
Dim oldPhotoName   : oldPhotoName   = UPLOAD__FORM("oldPhotoName")

Dim payMethodTxt

If payMethod = "SC0010" Then 
	payMethodTxt = "ī�����"
ElseIf payMethod = "SC0030" Then 
	payMethodTxt = "�ǽð� ������ü"
ElseIf payMethod = "SC0060" Then 
	payMethodTxt = "�ڵ�������"
ElseIf payMethod = "SC0040" Then 
	payMethodTxt = "��������Ա�"
End If

If PhotoName <>"" Then 
	If FILE_CHECK_EXT_JPG(PhotoName) = True Then
'		If 0 = UPLOAD__FORM("PhotoName").FileLen Then 
'			With Response
'			 .Write "<script type='text/javascript'>alert('�߸��� �����Դϴ�. [JPG,JPEG] ���ϸ� ������ּ���.');window.opener.location.reload();window.close()</script>"
'			 .End
'			End With
'		End If
		If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("PhotoName").FileLen Then 
			PhotoName = DextFileUpload("PhotoName",UPLOAD_BASE_PATH & savePath,0)
		Else
			With Response
			 .Write "<script type='text/javascript'>alert('������ ũ��� 20MB �� �ѱ�� �����ϴ�.');window.opener.location.reload();window.close()</script>"
			 .End
			End With
		End If
	Else

		'With Response
		 '.Write "<script type='text/javascript'>alert('�߸��� �����Դϴ�. [JPG,JPEG,GIF,PNG] ���ϸ� ������ּ���.');window.opener.location.reload();window.close()</script>"
		'.End
		'End With

	End If
	If oldPhotoName <> "" Then
		Set FSO = CreateObject("Scripting.FileSystemObject")
			If (FSO.FileExists(UPLOAD_BASE_PATH & savePath & oldPhotoName)) Then	' ���� �̸��� ������ ���� �� ����
				fso.deletefile(UPLOAD_BASE_PATH & savePath & oldPhotoName)
			End If
		set FSO = Nothing
	End If
Else
	PhotoName = oldPhotoName
End If

'If PhotoName = "" Then 
'	With Response
'	 .Write "<script type='text/javascript'>alert('�߸��� �����Դϴ�. [JPG,JPEG,GIF,PNG] ���ϸ� ������ּ���.');window.opener.location.reload();window.close()</script>"
'	 .End
'	End With
'End If

Sub InsertPhoto()
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_

	"UPDATE [dbo].[SP_USER_MEMBER] SET " &_
	"	 [FirstName] = ? " &_
    "	,[LastName]  = ? " &_
    "	,[Photo]     = ? " &_
	"WHERE [UserIdx] = ? "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@FirstName" ,adVarChar , adParamInput, 50  , FirstName )
		.Parameters.Append .CreateParameter( "@LastName"  ,adVarChar , adParamInput, 50  , LastName )
		.Parameters.Append .CreateParameter( "@Photo"     ,adVarChar , adParamInput, 200 , PhotoName )
		.Parameters.Append .CreateParameter( "@UserIdx"   ,adInteger , adParamInput, 0   , session("UserIdx") )
		.Execute
	End with
	call cmdclose()
End Sub


Call Expires()
Call dbopen()
	Call getView()
	Call InsertPhoto()
Call dbclose()







if FI_Idx = "" or IsNull( FI_Idx ) Or areaIdx = "" or IsNull( areaIdx ) Or USER_UserIdx = "" or IsNull( USER_UserIdx ) Then
	With Response
	 .Write "<script type='text/javascript'>alert('�߸��� ����Դϴ�.');window.close()</script>"
	 .End
	End With
end If

' �ߺ�
If FI_CntDuplicate > 0 Then 
	With Response
	 .Write "<script type='text/javascript'>alert('�̹� ��ϵ� �������� �Դϴ�.\n\n�������������� Ȯ���� �ּ���.');window.close()</script>"
	 .End
	End With
End If
' �������� �հݿ���
'Response.write FI_CntDuplicate_program
If FI_CntDuplicate_program > 0 Then 
	With Response
	 .Write "<script type='text/javascript'>alert('�̹� �հ��� �ڰ����� �Դϴ�.\n\n�������������� Ȯ���� �ּ���.');window.close()</script>"
	 .End
	End With
End If
' ����
If FI_CK_EndDate < Left(Now(),10) Then
	With Response
	 .Write "<script type='text/javascript'>alert('���� �����Ǿ����ϴ�.');window.close()</script>"
	 .End
	End With
End If
' ������
If FI_CK_StartDate > Left(Now(),10) Then 
	With Response
	 .Write "<script type='text/javascript'>alert('���� �����Ⱓ�� �ƴմϴ�.');window.close()</script>"
	 .End
	End With
End If
' �ο�����
If int(FI_CK_MaxNumber) <= int(FI_CK_CNT_APP) Then 
	With Response
	 .Write "<script type='text/javascript'>alert('���� �����ʰ�!');window.close()</script>"
	 .End
	End With
End If

PrograName = FI_Name

If FI_Class = "1" Then
	PrograName = PrograName & " 1��"
ElseIf FI_Class = "2" Then
	PrograName = PrograName & " 2��"
ElseIf FI_Class = "3" Then
	PrograName = PrograName & " 3��"
End If

If FI_Kind = "1" Then
	PrograName = PrograName & " [�ʱ�]"
ElseIf FI_Kind = "2" Then
	PrograName = PrograName & " [�Ǳ�]"
End If

Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Idx INT,@UserIdx INT ; " &_
	"SET @Idx = ?; " &_
	"SET @UserIdx = ?; " &_

	"DECLARE @CntDuplicate INT , @CK_StartDate DATETIME , @CK_EndDate DATETIME , @CK_MaxNumber INT , @CK_CNT_APP INT ;" &_
	"SET @CntDuplicate = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM_APP] WHERE [ProgramIdx] = @Idx AND [UserIdx] = @UserIdx AND [State] != 2 )  " &_

	"DECLARE @CntDuplicate_program INT " &_
	"DECLARE @T TABLE(IDX INT) " &_
	"INSERT INTO @T(IDX)" &_
	"	select A.[Idx] from [dbo].[SP_PROGRAM] A" &_
	"	INNER JOIN [dbo].[SP_PROGRAM] B" &_
	"	on(A.[CodeIdx] = B.[CodeIdx] AND A.[Kind] = B.[Kind] AND A.[Class] = B.[Class])" &_
	"	where B.[Idx] = @Idx " &_

	"SET @CntDuplicate_program = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM_APP] WHERE [ProgramIdx] IN( select [IDX] FROM @T ) AND [UserIdx] = @UserIdx AND [State] = 10 )  " &_
	
	"SELECT " &_
	"	 @CK_StartDate = CONVERT(varchar(10),A.[StartDate],23) " &_
	"	,@CK_EndDate   = CONVERT(varchar(10),A.[EndDate],23) " &_
	"	,@CK_MaxNumber = ISNULL( A.[MaxNumber],0 ) " &_
	"	,@CK_CNT_APP   = ISNULL(B.[CNT_APP],0) " &_
	"FROM [dbo].[SP_PROGRAM] A " &_
	"LEFT JOIN ( " &_
	"	SELECT " &_
	"		 [ProgramIdx] " &_
	"		,COUNT(*) AS [CNT_APP] " &_
	"	FROM [dbo].[SP_PROGRAM_APP] " &_
	"	WHERE [State] != 2 " &_
	"	GROUP BY [ProgramIdx] " &_
	") B ON(A.[Idx] = B.[ProgramIdx] ) " &_
	"WHERE [Dellfg] = 0 " &_
	"AND A.[Idx] = @Idx " &_

	"SELECT " &_
	"	 @CntDuplicate AS [CntDuplicate] " &_
	"	,@CntDuplicate_program AS [CntDuplicate_program] " &_
	"	,@CK_StartDate AS [CK_StartDate] " &_
	"	,@CK_EndDate AS [CK_EndDate] " &_
	"	,@CK_MaxNumber AS [CK_MaxNumber] " &_
	"	,@CK_CNT_APP AS [CK_CNT_APP] " &_
	"	,A.[Idx]" &_
	"	,A.[Pay]" &_
	"	,A.[OnData]" &_
	"	,A.[Kind] " &_
	"	,A.[Class] " &_
	"	,B.[Name] " &_
	"FROM [dbo].[SP_PROGRAM] A " &_
	"INNER JOIN [dbo].[SP_COMM_CODE2] B ON(A.[CodeIdx] = B.[Idx]) " &_
	"WHERE A.[Idx] =  @Idx AND A.[Dellfg] = 0  " &_

	"SELECT " &_
	"	 [UserIdx]" &_
	"	,[UserName]" &_
	"	,[UserId]" &_
	"	,[UserBirth]" &_
	"	,[UserHphone1]" &_
	"	,[UserHphone2]" &_
	"	,[UserHphone3]" &_
	"	,[UserEmail]" &_
	"	,[UserAddr1]" &_
	"	,[UserAddr2]" &_
	"	,[Photo]" &_
	"	,[LastName]" &_
	"	,[FirstName]" &_
	"FROM [dbo].[SP_USER_MEMBER] WHERE [UserIdx] = @UserIdx AND [UserDelFg] = 0 "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@Idx" ,adInteger , adParamInput , 0 , programIdx )
		.Parameters.Append .CreateParameter( "@UserIdx" ,adInteger , adParamInput , 0 , Session("UserIdx") )
		set objRs = .Execute
	End with
	call cmdclose()
	'���α׷�����
	CALL setFieldValue(objRs, "FI")
	'ȸ������
	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "USER")

	Set objRs = Nothing
End Sub

'/*
' * [���� ������û ������(STEP2-1)]
' *
' * ���������������� �⺻ �Ķ���͸� ���õǾ� ������, ������ �ʿ��Ͻ� �Ķ���ʹ� �����޴����� �����Ͻþ� �߰� �Ͻñ� �ٶ��ϴ�.
' */

'/*
' * 1. �⺻���� ������û ���� ����
' *
' * �⺻������ �����Ͽ� �ֽñ� �ٶ��ϴ�.(�Ķ���� ���޽� POST�� ����ϼ���)
' */

CST_PLATFORM               = "service"					'LG���÷��� ���� ���� ����(test:�׽�Ʈ, service:����)
CST_MID                    = "soribiblue"            '�������̵�(LG���÷������� ���� �߱޹����� �������̵� �Է��ϼ���)
																 '�׽�Ʈ ���̵�� 't'�� �ݵ�� �����ϰ� �Է��ϼ���.
if CST_PLATFORM = "test" then                                    '�������̵�(�ڵ�����)
	LGD_MID = "t" & CST_MID
else
	LGD_MID = CST_MID
end If
oid = replace(date,"-","") & Hour(now) & Minute(now) & second(now) & Session("UserIdx")

LGD_OID                    = oid                                '�ֹ���ȣ(�������� ����ũ�� �ֹ���ȣ�� �Է��ϼ���)
LGD_AMOUNT                 = FI_Pay								'�����ݾ�("," �� ������ �����ݾ��� �Է��ϼ���)
LGD_MERTKEY                = "15389449198e287da43d4785c51b9ca1" '[�ݵ�� ����]����MertKey(mertkey�� ���������� -> ������� -> ���������������� Ȯ���ϽǼ� �ֽ��ϴ�')
LGD_BUYER                  = trim(Left(USER_UserName,10))       '�����ڸ�
LGD_PRODUCTINFO            = trim(PrograName)                   '��ǰ��
LGD_BUYEREMAIL             = trim(USER_UserEmail)               '������ �̸���
LGD_TIMESTAMP              = year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2) 'Ÿ�ӽ�����
LGD_CUSTOM_SKIN            = "red"                               '�������� ����â ��Ų (red, blue, cyan, green, yellow)
LGD_BUYERID          	   = trim(Left(USER_UserId,15))          '������ ���̵�
LGD_BUYERIP          	   = trim(g_uip)      	                 '������IP
LGD_CUSTOM_USABLEPAY	   = payMethod                           '"SC0010-SC0030-SC0060"             '�������� 
LGD_BUYERPHONE			   = Trim(USER_UserHphone1) &"-"&Trim(USER_UserHphone2) &"-"&Trim(USER_UserHphone3)             '��ȭ��ȣ
LGD_DISPLAY_BUYERPHONE	   = "Y"								 '�޴�����ȣ�Է¿���
LGD_DISPLAY_BUYEREMAIL	   = "Y"								 '�̸����ּ��Է¿���
LGD_AUTOFILLYN_BUYER	   = "Y"								 '�����ڸ� �ڵ�ä��
LGD_CASHRECEIPTYN		   = "N"								 '���ݿ������߱� ��뿩��
LGD_INSTALLRANGE           = "0"								 '�Һ� 0:2:3:4:5:6:7:8:9:10:11:12

LGD_DISABLECARD            = ""
LGD_DISPLAY_ACCOUNTPID     = "N"


'/*
' * �������(������) ���� ������ �Ͻô� ��� �Ʒ� LGD_CASNOTEURL �� �����Ͽ� �ֽñ� �ٶ��ϴ�.
' */
LGD_CASNOTEURL             = "http://" & Request.ServerVariables("SERVER_NAME") & "/LGU_XPay_ASP/cas_noteurl.asp"

'/*
' *************************************************
' * 2. MD5 �ؽ���ȣȭ (�������� ������) - BEGIN
' *
' * MD5 �ؽ���ȣȭ�� �ŷ� �������� �������� ����Դϴ�.
' *************************************************
' *
' * �ؽ� ��ȣȭ ����( LGD_MID + LGD_OID + LGD_AMOUNT + LGD_TIMESTAMP + LGD_MERTKEY )
' * LGD_MID          : �������̵�
' * LGD_OID          : �ֹ���ȣ
' * LGD_AMOUNT       : �ݾ�
' * LGD_TIMESTAMP    : Ÿ�ӽ�����
' * LGD_MERTKEY      : ����MertKey (mertkey�� ���������� -> ������� -> ���������������� Ȯ���ϽǼ� �ֽ��ϴ�)
' *
' * MD5 �ؽ������� ��ȣȭ ������ ����
' * LG���÷������� �߱��� ����Ű(MertKey)�� ȯ�漳�� ����(lgdacom/conf/mall.conf)�� �ݵ�� �Է��Ͽ� �ֽñ� �ٶ��ϴ�.
' */
LGD_HASHDATA = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_TIMESTAMP & LGD_MERTKEY )
LGD_CUSTOM_PROCESSTYPE = "TWOTR"
'/*
' *************************************************
' * 2. MD5 �ؽ���ȣȭ (�������� ������) - END
' *************************************************
' */
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=EUC-KR">
<TITLE> �ѱ��ܽ�������ȸ ����������  </TITLE>

<script type="text/javascript">
<!--

window.onload = function(){
	isActiveXOK();
	var innerBody=document.body
	var innerHeight = document.compatMode == "CSS1Compat" ?
					   document.documentElement.scrollHeight : document.body.scrollHeight;
	resizeTo(460,innerHeight+52)
}

/*
 * �������� ������û�� PAYKEY�� �޾Ƽ� �������� ��û.
 */
function doPay_ActiveX(){
	document.getElementById('LGD_BUTTON2').innerHTML='���� ��û�� �Դϴ�.';
    ret = xpay_check(document.getElementById('LGD_PAYINFO'), '<%= CST_PLATFORM %>');

    if (ret=="00"){     //ActiveX �ε� ����
        var LGD_RESPCODE        = dpop.getData('LGD_RESPCODE');       //����ڵ�
        var LGD_RESPMSG         = dpop.getData('LGD_RESPMSG');        //����޼���

        if( "0000" == LGD_RESPCODE ) { //��������
            var LGD_PAYKEY      = dpop.getData('LGD_PAYKEY');         //LG���÷��� ����KEY
            var msg = "������� : " + LGD_RESPMSG + "\n";
            msg += "LGD_PAYKEY : " + LGD_PAYKEY +"\n\n";
            document.getElementById('LGD_PAYKEY').value = LGD_PAYKEY;
            //alert(msg);
            document.getElementById('LGD_PAYINFO').submit();
        } else { //��������
            alert("������ �����Ͽ����ϴ�. " + LGD_RESPMSG);
            /*
             * �������� ȭ�� ó��
             */
			 document.getElementById('LGD_BUTTON2').innerHTML='<img src="./img/btn.gif" onclick="doPay_ActiveX();" style="cursor:pointer;">';
        }
    } else {
        alert("LG U+ ���ڰ����� ���� ActiveX Control��  ��ġ���� �ʾҽ��ϴ�.");
        /*
         * �������� ȭ�� ó��
         */
		 document.getElementById('LGD_BUTTON2').innerHTML='<img src="./img/btn.gif" onclick="doPay_ActiveX();" style="cursor:pointer;">';
    }
}

function isActiveXOK(){
	if(lgdacom_atx_flag == true){
    	document.getElementById('LGD_BUTTON1').style.display='none';
        document.getElementById('LGD_BUTTON2').style.display='';
	}else{
		document.getElementById('LGD_BUTTON1').style.display='';
        document.getElementById('LGD_BUTTON2').style.display='none';	
	}
}
-->
</script>
<style type="text/css">
td{font-size:12px;}
.td_title{width:95px;padding-left:10px;}
.td_cont_pink{color:#ff469d}
</style>
</head>

<body style="margin:0px;">
<div id="LGD_ACTIVEX_DIV"/> <!-- ActiveX ��ġ �ȳ� Layer �Դϴ�. �������� ������. -->
<form method="post" id="LGD_PAYINFO" action="payres.asp">


<table cellpadding=0 cellspacing=0 width="450">
	<tr>
		<td><img src="./img/title.gif"></td>
	</tr>
	<tr>
		<td align=center>
			<table cellpadding=0 cellspacing=0 width="100%" align=center bgcolor="#ffffff">
				<tr>
					<td align=center>
						<table cellpadding=0 cellspacing=0 border=0 align=center width=370>
							<tr>
								<td colspan=2 style="padding:20px 0px 20px 0px;"><img src="./img/sub_title.gif"></td>
							</tr>
							<tr height="28">
								<td class=td_title><img src="./img/icon_arrow.gif"> &nbsp��ǰ��</td>
								<td width=275><%=LGD_PRODUCTINFO%></td>
							</tr>
							<tr height="28">
								<td class=td_title><img src="./img/icon_arrow.gif"> &nbsp��������</td>
								<td><%=payMethodTxt%></td>
							</tr>
							<tr height="28">
								<td class=td_title><img src="./img/icon_arrow.gif"> &nbsp�����ݾ�</td>
								<td class=td_cont_pink><%= FormatNumber(LGD_AMOUNT,0) %> (VAT����)</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>

		</td>
	</tr>
	<tr>
		<td style="padding-top:30px;" align=center><img src="./img/bar.gif"></td>
	</tr>
	<tr>
		<td align=center height=70>
			<div id="LGD_BUTTON1">������ ���� ����� �ٿ� ���̰ų�, ����� ��ġ���� �ʾҽ��ϴ�. </div>
			<div id="LGD_BUTTON2" style="display:none"><img src="./img/btn.gif" onclick="doPay_ActiveX();" style="cursor:pointer;"></div>
		</td>
	</tr>
	<tr>
		<td align=center style="padding:0px 0px 20px 0px;">�� �������� ������ ������ TEL : <span class=td_cont_pink><b>1800-6288</b></span> ������ �����ֽñ� �ٶ��ϴ�.</td>
	</tr>
	<tr>
		<td height="18" bgcolor="617bff">&nbsp;</td>
	</tr>


</table>
<br>

<br>
<input type="hidden" name="CST_PLATFORM"                value="<%= CST_PLATFORM %>">                   <!-- �׽�Ʈ, ���� ���� -->
<input type="hidden" name="CST_MID"                     value="<%= CST_MID %>">                        <!-- �������̵� -->
<input type="hidden" name="LGD_MID"                     value="<%= LGD_MID %>">                        <!-- �������̵� -->
<input type="hidden" name="LGD_OID"                     value="<%= LGD_OID %>">                        <!-- �ֹ���ȣ -->
<input type="hidden" name="LGD_BUYER"                   value="<%= LGD_BUYER %>">                      <!-- ������ -->
<input type="hidden" name="LGD_PRODUCTINFO"             value="<%= LGD_PRODUCTINFO %>">                <!-- ��ǰ���� -->
<input type="hidden" name="LGD_AMOUNT"                  value="<%= LGD_AMOUNT %>">                     <!-- �����ݾ� -->
<input type="hidden" name="LGD_BUYEREMAIL"              value="<%= LGD_BUYEREMAIL %>">                 <!-- ������ �̸��� -->
<input type="hidden" name="LGD_CUSTOM_SKIN"             value="<%= LGD_CUSTOM_SKIN %>">                <!-- ����â SKIN -->
<input type="hidden" name="LGD_CUSTOM_PROCESSTYPE"      value="<%= LGD_CUSTOM_PROCESSTYPE %>">         <!-- Ʈ����� ó����� -->
<input type="hidden" name="LGD_TIMESTAMP"               value="<%= LGD_TIMESTAMP %>">                  <!-- Ÿ�ӽ����� -->
<input type="hidden" name="LGD_HASHDATA"                value="<%= LGD_HASHDATA %>">                   <!-- MD5 �ؽ���ȣ�� -->
<input type="hidden" name="LGD_PAYKEY"                  id="LGD_PAYKEY">                               <!-- LG���÷��� PAYKEY(������ �ڵ�����)-->
<input type="hidden" name="LGD_VERSION"         		value="ASP_XPay_1.0">						   <!-- �������� (�������� ������) -->
<input type="hidden" name="LGD_BUYERIP"                 value="<%= LGD_BUYERIP %>">        			   <!-- ������IP -->
<input type="hidden" name="LGD_BUYERID"                 value="<%= LGD_BUYERID %>">           		   <!-- ������ID -->
<input type="hidden" name="LGD_CUSTOM_USABLEPAY"		value="<%= LGD_CUSTOM_USABLEPAY %>">           <!-- �������ǰ������ɼ��� -->
<input type="hidden" name="LGD_BUYERSSN"				value="<%= LGD_BUYERSSN %>">				   <!-- �������ֹι�ȣ -->
<input type="hidden" name="LGD_CHECKSSNYN"				value="<%= LGD_CHECKSSNYN %>">				   <!-- �������ֹι�ȣ üũ���� -->

<input type="hidden" name="LGD_BUYERPHONE"				value="<%= LGD_BUYERPHONE %>">				   <!-- ��������ȭ��ȣ -->
<input type="hidden" name="LGD_DISPLAY_BUYERPHONE"		value="<%= LGD_DISPLAY_BUYERPHONE %>">		   <!-- �޴�����ȣ�Է¿��� -->
<input type="hidden" name="LGD_DISPLAY_BUYEREMAIL"		value="<%= LGD_DISPLAY_BUYEREMAIL %>">		   <!-- �̸����ּ��Է¿��� -->
<input type="hidden" name="LGD_DISPLAY_ACCOUNTPID"		value="<%= LGD_DISPLAY_ACCOUNTPID %>">		   <!-- ��������ֹι�ȣ�Է¿��� -->
<input type="hidden" name="LGD_AUTOFILLYN_BUYER"		value="<%= LGD_AUTOFILLYN_BUYER %>">		   <!-- �����ڸ� �ڵ�ä�� -->
<input type="hidden" name="LGD_AUTOFILLYN_BUYERSSN"		value="<%= LGD_AUTOFILLYN_BUYERSSN %>">		   <!-- ������ �ֹι�ȣ �ڵ�ä�� -->
<input type="hidden" name="LGD_CASHRECEIPTYN"			value="<%= LGD_CASHRECEIPTYN %>">			   <!-- ���ݿ������߱޻�뿩�� -->
<input type="hidden" name="LGD_DISABLECARD"				value="<%= LGD_DISABLECARD %>">                <!-- ���Ұ���ī�� -->
<input type="hidden" name="LGD_INSTALLRANGE"			value="<%= LGD_INSTALLRANGE %>">                <!-- ǥ���Һΰ����� -->

<input type="hidden" name="programIdx" value="<%= programIdx %>">     <!-- ���α׷� IDX -->
<input type="hidden" name="areaIdx"    value="<%= areaIdx %>">        <!-- ������ IDX -->

<!-- �������(������) ���������� �Ͻô� ���  �Ҵ�/�Ա� ����� �뺸�ޱ� ���� �ݵ�� LGD_CASNOTEURL ������ LG ���÷����� �����ؾ� �մϴ� . -->
<input type="hidden" name="LGD_CASNOTEURL"           value="<%= LGD_CASNOTEURL %>">                 <!-- ������� NOTEURL -->


</form>
</body>
<!--  xpay.js�� �ݵ��  body �ؿ� �νñ� �ٶ��ϴ�. -->
<!--  UTF-8 ���ڵ� ��� �ô� xpay.js ��� xpay_utf-8.js ��  ȣ���Ͻñ� �ٶ��ϴ�.-->
<%
     protocol = "http"
     If request.serverVariables("SERVER_PORT") = "443" Then protocol = "https"

     if CST_PLATFORM = "test" then
     	port = "7080"
     	If request.serverVariables("SERVER_PORT") = "443" Then port = "7443"
        Response.Write "<script language='javascript' src='"& protocol &"://xpay.uplus.co.kr:" & port & "/xpay/js/xpay.js' type='text/javascript'>"
     else
        Response.Write "<script language='javascript' src='"& protocol &"://xpay.uplus.co.kr/xpay/js/xpay.js' type='text/javascript'>"
     end if
%>
</script>
</html>