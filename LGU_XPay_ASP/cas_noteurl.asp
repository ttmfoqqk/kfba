<!-- #include file = "../_lib/header.asp" -->
<!-- #include file="./lgdacom/md5.asp" -->
<%
Call Expires()
Call dbopen()
'/*
' * [���� �������ó��(DB) ������]
' *
' * 1) ������ ������ ���� hashdata�� ������ �ݵ�� �����ϼž� �մϴ�.
' *
' */
LGD_RESPCODE            = trim(request("LGD_RESPCODE"))             '// �����ڵ�: 0000(����) �׿� ����
LGD_RESPMSG             = trim(request("LGD_RESPMSG"))              '// ����޼���
LGD_MID                 = trim(request("LGD_MID"))                  '// �������̵�
LGD_OID                 = trim(request("LGD_OID"))                  '// �ֹ���ȣ
LGD_AMOUNT              = trim(request("LGD_AMOUNT"))               '// �ŷ��ݾ�
LGD_TID                 = trim(request("LGD_TID"))                  '// LG���÷������� �ο��� �ŷ���ȣ
LGD_PAYTYPE             = trim(request("LGD_PAYTYPE"))              '// ���������ڵ�
LGD_PAYDATE             = trim(request("LGD_PAYDATE"))              '// �ŷ��Ͻ�(�����Ͻ�/��ü�Ͻ�)
LGD_HASHDATA            = trim(request("LGD_HASHDATA"))             '// �ؽ���
LGD_FINANCECODE         = trim(request("LGD_FINANCECODE"))          '// ��������ڵ�(�����ڵ�)
LGD_FINANCENAME         = trim(request("LGD_FINANCENAME"))          '// ��������̸�(�����̸�)
LGD_ESCROWYN            = trim(request("LGD_ESCROWYN"))             '// ����ũ�� ���뿩��
LGD_TIMESTAMP           = trim(request("LGD_TIMESTAMP"))            '// Ÿ�ӽ�����
LGD_ACCOUNTNUM          = trim(request("LGD_ACCOUNTNUM"))           '// ���¹�ȣ(�������Ա�)
LGD_CASTAMOUNT          = trim(request("LGD_CASTAMOUNT"))           '// �Ա��Ѿ�(�������Ա�)
LGD_CASCAMOUNT          = trim(request("LGD_CASCAMOUNT"))           '// ���Աݾ�(�������Ա�)
LGD_CASFLAG             = trim(request("LGD_CASFLAG"))              '// �������Ա� �÷���(�������Ա�) - 'R':�����Ҵ�, 'I':�Ա�, 'C':�Ա����
LGD_CASSEQNO            = trim(request("LGD_CASSEQNO"))             '// �Աݼ���(�������Ա�)
LGD_CASHRECEIPTNUM      = trim(request("LGD_CASHRECEIPTNUM"))       '// ���ݿ����� ���ι�ȣ
LGD_CASHRECEIPTSELFYN   = trim(request("LGD_CASHRECEIPTSELFYN"))    '// ���ݿ����������߱������� Y: �����߱��� ����, �׿� : ������
LGD_CASHRECEIPTKIND     = trim(request("LGD_CASHRECEIPTKIND"))      '// ���ݿ����� ���� 0: �ҵ������ , 1: ����������
LGD_PAYER            	= trim(request("LGD_PAYER"))             	'// �Ա��ڸ�

'/*
' * ��������
' */
LGD_BUYER               = trim(request("LGD_BUYER"))                '// ������
LGD_PRODUCTINFO         = trim(request("LGD_PRODUCTINFO"))          '// ��ǰ��
LGD_BUYERID             = trim(request("LGD_BUYERID"))              '// ������ ID
LGD_BUYERADDRESS        = trim(request("LGD_BUYERADDRESS"))         '// ������ �ּ�
LGD_BUYERPHONE          = trim(request("LGD_BUYERPHONE"))           '// ������ ��ȭ��ȣ
LGD_BUYEREMAIL          = trim(request("LGD_BUYEREMAIL"))           '// ������ �̸���
LGD_BUYERSSN            = trim(request("LGD_BUYERSSN"))             '// ������ �ֹι�ȣ
LGD_PRODUCTCODE         = trim(request("LGD_PRODUCTCODE"))          '// ��ǰ�ڵ�
LGD_RECEIVER            = trim(request("LGD_RECEIVER"))             '// ������
LGD_RECEIVERPHONE       = trim(request("LGD_RECEIVERPHONE"))        '// ������ ��ȭ��ȣ
LGD_DELIVERYINFO        = trim(request("LGD_DELIVERYINFO"))         '// �����

'/*
' * hashdata ������ ���� mertkey�� ���������� -> ������� -> ���������������� Ȯ���ϽǼ� �ֽ��ϴ�.
' * LG���÷������� �߱��� ����Ű�� �ݵ�� ������ �ֽñ� �ٶ��ϴ�.
' */
LGD_MERTKEY = "15389449198e287da43d4785c51b9ca1"  '//mertkey
LGD_HASHDATA2 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )
'LGD_HASHDATA2 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )

'/*
' * ���� ó����� ���ϸ޼���
' *
' * OK  : ���� ó����� ����
' * �׿� : ���� ó����� ����
' *
' * �� ���ǻ��� : ������ 'OK' �����̿��� �ٸ����ڿ��� ���ԵǸ� ����ó�� �ǿ��� �����Ͻñ� �ٶ��ϴ�.
' */
resultMSG = "������� ���� DBó��(LGD_CASNOTEURL) ������� �Է��� �ֽñ� �ٶ��ϴ�."

if LGD_HASHDATA2 = LGD_HASHDATA then
	'//�ؽ��� ������ �����̸�
	if LGD_RESPCODE = "0000" then
		'//������ �����̸�
		if LGD_CASFLAG = "R" then
			'/*
			' * ������ �Ҵ� ���� ��� ���� ó��(DB) �κ�
			' * ���� ��� ó���� �����̸� "OK"
			' */
			resultMSG = "OK"
		elseif LGD_CASFLAG = "I" then
			'/*
			' * ������ �Ա� ���� ��� ���� ó��(DB) �κ�
			' * ���� ��� ó���� �����̸� "OK"
			' */
			
			Call getView()
			
			
			Dim isDBOK  : isDBOK  = true
			Dim sNumber : sNumber = ""
			Dim State   : State   = 1

			'���˹�ȣ ����
			'���ó⵵2�ڸ� + ���ÿ�2�ڸ� + ������3�ڸ� + ��������1�ڸ� + �ʱ�/�Ǳ�1�ڸ� + �޼� 1�ڸ� + ��Ϲ�ȣ3�ڸ�
			Dim sNumber1 : sNumber1 = Mid(FI_OnData,3,2)
			Dim sNumber2 : sNumber2 = Mid(FI_OnData,6,2)
			Dim sNumber3 : sNumber3 = FI_AreaCode
			Dim sNumber4 : sNumber4 = FI_ProgramCode
			Dim sNumber5 : sNumber5 = FI_Kind
			Dim sNumber6 : sNumber6 = FI_Class
			Dim sNumber7 : sNumber7 = lpad( FI_AppCode , "0" , 3 )

			sNumber = sNumber1 & sNumber2 & sNumber3 & sNumber4 & sNumber5 & sNumber6 & sNumber7

			If LGD_CASTAMOUNT = LGD_AMOUNT Then 
			State = 0
			End If

			'������ ���˹�ȣ 13�ڸ� üũ
			if Len(sNumber) <> 13 Then
				isDBOK = False
			End If

			If isDBOK = True Then 
				Call UpdateI()
				If RESULT_ERR = "0" Then 
					resultMSG = "OK"
				End If
			End If

		elseif LGD_CASFLAG = "C" then
			'/*
			' * ������ �Ա���� ���� ��� ���� ó��(DB) �κ�
			' * ���� ��� ó���� �����̸� "OK"
			' */
			Call getView()

			Call UpdateC()
			If RESULT_ERR = "0" Then 
				resultMSG = "OK"
			End If

			
		end if
	else
		'//������ �����̸�
		'/*
		' * �ŷ����� ��� ���� ó��(DB) �κ�
		' * ������� ó���� �����̸� "OK"
		' */
		resultMSG = "OK"
	end if
else
	'//�ؽ����� ������ �����̸�
	'/*
	' * hashdata���� ���� �α׸� ó���Ͻñ� �ٶ��ϴ�.
	' */
	resultMSG = "������� ���� DBó��(LGD_CASNOTEURL) �ؽ��� ������ �����Ͽ����ϴ�."
end If

Call dbclose()

Response.Write(resultMSG)






Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @LGD_TID VARCHAR(24),@LGD_OID VARCHAR(64) , @LGD_IDX INT , @PROGRAM_IDX INT , @AREA_IDX INT , @APP_IDX INT;" &_
	"SET @LGD_TID = ?; " &_
	"SET @LGD_OID = ?; " &_
	"SET @LGD_IDX = (SELECT [Idx] FROM [dbo].[SP_Pay_LGD] WHERE [LGD_OID] = @LGD_OID ); " &_
	
	"SELECT " &_
	"	 @APP_IDX     = [Idx] " &_
	"	,@PROGRAM_IDX = [ProgramIdx] " &_
	"	,@AREA_IDX    = [AreaIdx] " &_
	"FROM [dbo].[SP_PROGRAM_APP] " &_
	"WHERE [LgdIdx] = @LGD_IDX " &_

	"SELECT " &_
	"	 [Idx]" &_
	"	,[Pay]" &_
	"	,convert(varchar(10),[OnData],23) AS [OnData] " &_
	"	,( SELECT COUNT(*) + 1 FROM [dbo].[SP_COMM_CODE2] where [PIdx] = 17 and [idx] < [CodeIdx] ) AS [ProgramCode] " &_
	"	,( SELECT ISNULL([Code],'000') FROM [dbo].[SP_PROGRAM_AREA] where [Idx] = @AREA_IDX ) AS [AreaCode] " &_
	"	,( SELECT COUNT(*) + 1 FROM [dbo].[SP_PROGRAM_APP] where [ProgramIdx] = @PROGRAM_IDX AND [AreaIdx] = @AREA_IDX AND [Idx] < @APP_IDX ) AS [AppCode] " &_
	"	,[Kind] " &_
	"	,[Class] " &_
	"	,@APP_IDX AS [APP_IDX] " &_
	"	,@LGD_IDX AS [LGD_IDX] " &_
	"FROM [dbo].[SP_PROGRAM] " &_
	"WHERE [Idx] =  @PROGRAM_IDX AND [Dellfg] = 0 "

	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@LGD_TID" ,adVarChar , adParamInput , 24 , LGD_TID )
		.Parameters.Append .CreateParameter( "@LGD_OID" ,adVarChar , adParamInput , 64 , LGD_OID )
		set objRs = .Execute
	End with
	call cmdclose()
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub

Sub UpdateI()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @ERR int ;" &_
	"set @ERR = 0 " &_

	"DECLARE @LGD_IDX INT , @APP_IDX INT , @LGD_RESPMSG VARCHAR(512) , @LGD_CASTAMOUNT VARCHAR(12) , @LGD_CASCAMOUNT VARCHAR(12) , @LGD_CASFLAG VARCHAR(10) , @LGD_CASSEQNO VARCHAR(3) , @State INT , @Snumber VARCHAR(50);" &_

	"SET @LGD_IDX        = ?;" &_
	"SET @APP_IDX        = ?;" &_
	"SET @LGD_RESPMSG    = ?;" &_
	"SET @LGD_CASTAMOUNT = ?;" &_
	"SET @LGD_CASCAMOUNT = ?;" &_
	"SET @LGD_CASFLAG    = ?;" &_
	"SET @LGD_CASSEQNO   = ?;" &_
	"SET @State          = ?;" &_
	"SET @Snumber        = ?;" &_

	"BEGIN TRAN " &_

	"UPDATE [dbo].[SP_Pay_LGD] SET " &_
	"	 [LGD_RESPMSG]    = @LGD_RESPMSG " &_
	"	,[LGD_CASTAMOUNT] = @LGD_CASTAMOUNT " &_
	"	,[LGD_CASCAMOUNT] = @LGD_CASCAMOUNT " &_
	"	,[LGD_CASFLAG]    = @LGD_CASFLAG " &_
	"	,[LGD_CASSEQNO]   = @LGD_CASSEQNO " &_
	"WHERE [Idx] =  @LGD_IDX " &_

	"UPDATE [dbo].[SP_PROGRAM_APP] SET  " &_
	"	 [State]      = @State " &_
	"	,[Snumber]    = @Snumber " &_
	"	,[NocachDate] = CASE @State WHEN 0 THEN getdate() ELSE NULL END " &_
	"WHERE [Idx] = @APP_IDX " &_

	"IF @@error <> 0 " &_
    "BEGIN " &_
	"	ROLLBACK TRAN " &_
	"	set @ERR = 1 " &_
	"	RETURN " &_
    "END " &_
	"COMMIT TRAN " &_

	"select (@ERR) as ERR "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@LGD_IDX"        ,adInteger , adParamInput , 0   , FI_LGD_IDX )
		.Parameters.Append .CreateParameter( "@APP_IDX"        ,adInteger , adParamInput , 0   , FI_APP_IDX )
		.Parameters.Append .CreateParameter( "@LGD_RESPMSG"    ,adVarChar , adParamInput , 512 , LGD_RESPMSG )
		.Parameters.Append .CreateParameter( "@LGD_CASTAMOUNT" ,adVarChar , adParamInput , 12  , LGD_CASTAMOUNT )
		.Parameters.Append .CreateParameter( "@LGD_CASCAMOUNT" ,adVarChar , adParamInput , 12  , LGD_CASCAMOUNT )
		.Parameters.Append .CreateParameter( "@LGD_CASFLAG"    ,adVarChar , adParamInput , 10  , LGD_CASFLAG )
		.Parameters.Append .CreateParameter( "@LGD_CASSEQNO"   ,adVarChar , adParamInput , 3   , LGD_CASSEQNO )		
		.Parameters.Append .CreateParameter( "@State"          ,adInteger , adParamInput , 0   , State )
		.Parameters.Append .CreateParameter( "@Snumber"        ,adVarChar , adParamInput , 50  , sNumber )

		set objRs = .Execute
	End with
	call cmdclose()

	' �ʵ� �ε����� ���� ����.
	CALL setFieldValue(objRs, "RESULT")
	objRs.close	: Set objRs = Nothing
end Sub


Sub UpdateC()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @ERR int ;" &_
	"set @ERR = 0 " &_

	"DECLARE @LGD_IDX INT , @APP_IDX INT , @LGD_RESPMSG VARCHAR(512) , @LGD_CASTAMOUNT VARCHAR(12) , @LGD_CASCAMOUNT VARCHAR(12) , @LGD_CASFLAG VARCHAR(10) , @LGD_CASSEQNO VARCHAR(3) , @State INT , @Snumber VARCHAR(50);" &_

	"SET @LGD_IDX        = ?;" &_
	"SET @APP_IDX        = ?;" &_
	"SET @LGD_RESPMSG    = ?;" &_
	"SET @LGD_CASTAMOUNT = ?;" &_
	"SET @LGD_CASCAMOUNT = ?;" &_
	"SET @LGD_CASFLAG    = ?;" &_
	"SET @LGD_CASSEQNO   = ?;" &_
	"SET @State          = ?;" &_

	"BEGIN TRAN " &_

	"UPDATE [dbo].[SP_Pay_LGD] SET " &_
	"	 [LGD_RESPMSG]    = @LGD_RESPMSG " &_
	"	,[LGD_CASTAMOUNT] = @LGD_CASTAMOUNT " &_
	"	,[LGD_CASCAMOUNT] = @LGD_CASCAMOUNT " &_
	"	,[LGD_CASFLAG]    = @LGD_CASFLAG " &_
	"	,[LGD_CASSEQNO]   = @LGD_CASSEQNO " &_
	"WHERE [Idx] = @LGD_IDX " &_

	"UPDATE [dbo].[SP_PROGRAM_APP] SET  " &_
	"	[State] = @State " &_
	"WHERE [Idx] = @APP_IDX " &_

	"IF @@error <> 0 " &_
    "BEGIN " &_
	"	ROLLBACK TRAN " &_
	"	set @ERR = 1 " &_
	"	RETURN " &_
    "END " &_
	"COMMIT TRAN " &_

	"select (@ERR) as ERR "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@LGD_IDX"        ,adInteger , adParamInput , 0   , FI_LGD_IDX )
		.Parameters.Append .CreateParameter( "@APP_IDX"        ,adInteger , adParamInput , 0   , FI_APP_IDX )
		.Parameters.Append .CreateParameter( "@LGD_RESPMSG"    ,adVarChar , adParamInput , 512 , LGD_RESPMSG )
		.Parameters.Append .CreateParameter( "@LGD_CASTAMOUNT" ,adVarChar , adParamInput , 12  , LGD_CASTAMOUNT )
		.Parameters.Append .CreateParameter( "@LGD_CASCAMOUNT" ,adVarChar , adParamInput , 12  , LGD_CASCAMOUNT )
		.Parameters.Append .CreateParameter( "@LGD_CASFLAG"    ,adVarChar , adParamInput , 10  , LGD_CASFLAG )
		.Parameters.Append .CreateParameter( "@LGD_CASSEQNO"   ,adVarChar , adParamInput , 3   , LGD_CASSEQNO )		
		.Parameters.Append .CreateParameter( "@State"          ,adInteger , adParamInput , 0   , 2 )
		set objRs = .Execute
	End with
	call cmdclose()

	' �ʵ� �ε����� ���� ����.
	CALL setFieldValue(objRs, "RESULT")
	objRs.close	: Set objRs = Nothing
end Sub

%>