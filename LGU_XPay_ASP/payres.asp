<!-- #include file = "../_lib/header.asp" -->
<%
if session("UserIdx") = "" or IsNull(session("UserIdx"))=True Then
	With Response
	 .Write "<script type='text/javascript'>alert('�α����� �ʿ��մϴ�.');window.opener.location.reload();window.close()</script>"
	 .End
	End With
end If

Dim programIdx : programIdx = RequestSet("programIdx" , "POST" , 0)
Dim areaIdx    : areaIdx    = RequestSet("areaIdx"    , "POST" , 0)

'/*
' * [����������û ������(STEP2-2)]
' *
' * LG���÷������� ���� �������� LGD_PAYKEY(����Key)�� ������ ���� ������û.(�Ķ���� ���޽� POST�� ����ϼ���)
' */

'configPath = "C:/lgdacom"  'LG���÷������� ������ ȯ������("/conf/lgdacom.conf, /conf/mall.conf") ��ġ ����. 
configPath = server.mapPath("\LGU_XPay_ASP\lgdacom\")
'configPath = "F:/home/swid_soribiblue/www/LGU_XPay_ASP/lgdacom"
'configPath = "F:/HOME/swid_soribiblue/www/LGU_XPay_ASP/lgdacom"


'/*
' *************************************************
' * 1.�������� ��û - BEGIN
' *  (��, ���� �ݾ�üũ�� ���Ͻô� ��� �ݾ�üũ �κ� �ּ��� ���� �Ͻø� �˴ϴ�.)
' *************************************************
' */
CST_PLATFORM               = trim(request("CST_PLATFORM"))
CST_MID                    = trim(request("CST_MID"))
if CST_PLATFORM = "test" then
	LGD_MID = "t" & CST_MID
else
	LGD_MID = CST_MID
end if
LGD_PAYKEY                 = trim(request("LGD_PAYKEY"))

Dim xpay            '������û API ��ü
Dim amount_check    '�ݾ׺� ���
Dim i, j
Dim itemName

'�ش� API�� ����ϱ� ���� setup.exe �� ��ġ�ؾ� �մϴ�.
Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
xpay.Init configPath, CST_PLATFORM

xpay.Init_TX(LGD_MID)
xpay.Set "LGD_TXNAME", "PaymentByKey"
xpay.Set "LGD_PAYKEY", LGD_PAYKEY

'�ݾ��� üũ�Ͻñ� ���ϴ� ��� �Ʒ� �ּ��� Ǯ� �̿��Ͻʽÿ�.
'DB_AMOUNT = "DB�� ���ǿ��� ������ �ݾ�" 	'�ݵ�� �������� �Ұ����� ��(DB�� ����)���� �ݾ��� �������ʽÿ�.
'xpay.Set "LGD_AMOUNTCHECKYN", "Y"
'xpay.Set "LGD_AMOUNT", DB_AMOUNT
	
'/*
' *************************************************
' * 1.�������� ��û(�������� ������) - END
' *************************************************
' */

'/*
' * 2. �������� ��û ���ó��
' *
' * ���� ������û ��� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
' */

Call Expires()
Call dbopen()

	Call getView()
	
		
	if  xpay.TX() then
		'1)������� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
		'Response.Write("������û�� �Ϸ�Ǿ����ϴ�. <br>")
		'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
		'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

		'Response.Write("�ŷ���ȣ : " & xpay.Response("LGD_TID", 0) & "<br>")
		'Response.Write("�������̵� : " & xpay.Response("LGD_MID", 0) & "<br>")
		'Response.Write("�����ֹ���ȣ : " & xpay.Response("LGD_OID", 0) & "<br>")
		'Response.Write("�����ݾ� : " & xpay.Response("LGD_AMOUNT", 0) & "<br>")
		'Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
		'Response.Write("����޼��� : " & xpay.Response("LGD_RESPMSG", 0) & "<p>")

		'Response.Write("[������û ��� �Ķ����]<br>")

		'�Ʒ��� ������û ��� �Ķ���͸� ��� ��� �ݴϴ�.
		'Dim itemCount
		'Dim resCount
		'itemCount = xpay.resNameCount
		'resCount = xpay.resCount

		'For i = 0 To itemCount - 1
		'	itemName = xpay.ResponseName(i)
		'	Response.Write(itemName & "&nbsp: ")
		'	For j = 0 To resCount - 1
		'		Response.Write(xpay.Response(itemName, j) & "<br>")
		'	Next
		'Next
			
		'Response.Write("<p>")
		  
		if xpay.resCode = "0000" then
			'����������û ��� ���� DBó��
			'Response.Write("����������û ��� ���� DBó���Ͻñ� �ٶ��ϴ�." & "<br>")
												
			'����������û ��� ���� DBó�� ���н� Rollback ó��
			isDBOK = true 'DBó�� ���н� false�� ������ �ּ���.
			
			'�������� Ȯ��
			if FI_Idx = "" or IsNull( FI_Idx ) Or areaIdx = "" or IsNull( areaIdx ) Or USER_UserIdx = "" or IsNull( USER_UserIdx ) Then
				isDBOK = False
			End If
			
			Dim sNumber     : sNumber     = ""
			Dim State       : State       = 1
			Dim LGD_PAYTYPE : LGD_PAYTYPE = xpay.Response("LGD_PAYTYPE",0)
			If Trim(LGD_PAYTYPE) = "SC0040" Then '������ �Ա��϶� ���˹�ȣ ���� X
				sNumber = ""
				State   = 1
			Else
			
				'���˹�ȣ ����
				'���ó⵵2�ڸ� + ���ÿ�2�ڸ� + ������3�ڸ� + ��������1�ڸ� + �ʱ�/�Ǳ�1�ڸ� + �޼� 1�ڸ� + ��Ϲ�ȣ3�ڸ�
				Dim sNumber1 : sNumber1 = Mid(FI_OnData,3,2)
				Dim sNumber2 : sNumber2 = Mid(FI_OnData,6,2)
				Dim sNumber3 : sNumber3 = FI_AreaCode
				Dim sNumber4 : sNumber4 = FI_ProgramCode
				Dim sNumber5 : sNumber5 = FI_Kind
				Dim sNumber6 : sNumber6 = FI_Class
				'Dim sNumber7 : sNumber7 = lpad( FI_AppCode , "0" , 3 )

				sNumber = sNumber1 & sNumber2 & sNumber3 & sNumber4 & sNumber5 & sNumber6
				State   = 0
				
				'������ ���˹�ȣ 10�ڸ� üũ
				if Len(sNumber) <> 10 Then
					'isDBOK = False
				End If

			End If

			If isDBOK = True Then 
				Call Inert()
				If FI_ERR > 0 Then 
					isDBOK = False
				Else
					ResultCode = "0000"
					AlertMsg = "�����Ǿ����ϴ�."
				End If
			End If
				
			If isDBOK = False Then

				'Response.Write("<p>")
				xpay.Rollback("���� DBó�� ���з� ���Ͽ� Rollback ó�� [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
					
				'Response.Write("TX Rollback Response_code = " & xpay.resCode & "<br>")
				'Response.Write("TX Rollback Response_msg = " & xpay.resMsg & "<p>")
					
				if "0000" = xpay.resCode then
					'Response.Write("�ڵ���Ұ� ���������� �Ϸ� �Ǿ����ϴ�.<br>")
					ResultCode = "0001"
					AlertMsg = "ERR ["&xpay.Response("LGD_RESPCODE", 0)&"] : DBó�� ���� \n\n�ڵ���Ұ� ���������� �Ϸ� �Ǿ����ϴ�."
				else
					'Response.Write("�ڵ���Ұ� ���������� ó������ �ʾҽ��ϴ�.<br>")
					ResultCode = "0002"
					AlertMsg = "ERR ["&xpay.Response("LGD_RESPCODE", 0)&"] : DBó�� ���� \n\�ڵ���Ұ� ���������� ó������ �ʾҽ��ϴ�.\n\n�����ڿ��� �����ϼ���."
				end if
			end if            	
		else
			'����������û ��� ���� DBó��
			'Response.Write("����������û ��� ���� DBó���Ͻñ� �ٶ��ϴ�." & "<br>")
			ResultCode = "0003"
			AlertMsg = "ERR ["& xpay.resCode &"] : ������û ��� ���� [ "& xpay.resMsg &" ]"

			Call Inert_Fail()
		end if
	else
		'2)API ��û���� ȭ��ó��
		'Response.Write("������û�� �����Ͽ����ϴ�. <br>")
		'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
		'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
			
		'������û ��� ���� ���� DBó��
		'Response.Write("����������û ��� ���� DBó���Ͻñ� �ٶ��ϴ�." & "<br>")

		ResultCode = "0004"
		AlertMsg = "ERR ["& xpay.resCode &"] : ������û ��� ���� [ "& xpay.resMsg &" ]"

		Call Inert_Fail()

	end If

Call dbclose()


Sub getView()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	',( SELECT COUNT(*) + 1 FROM [dbo].[SP_PROGRAM_APP] where [ProgramIdx] = @Idx AND [AreaIdx] = @AreaIdx AND [Snumber] is not null ) AS [AppCode] 
	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @Idx INT,@AreaIdx INT,@UserIdx INT ;" &_
	"SET @Idx = ?; " &_
	"SET @AreaIdx = ?; " &_
	"SET @UserIdx = ?; " &_	

	"SELECT " &_
	"	 A.[Idx]" &_
	"	,A.[Pay]" &_
	"	,convert(varchar(10),A.[OnData],23) AS [OnData]" &_
	"	,B.[Name] " &_
	"	,( SELECT COUNT(*) + 1 FROM [dbo].[SP_COMM_CODE2] where [PIdx] = 17 and [idx] < A.[CodeIdx] ) AS [ProgramCode] " &_
	"	,( SELECT ISNULL([Code],'000') FROM [dbo].[SP_PROGRAM_AREA] where [Idx] = @AreaIdx ) AS [AreaCode] " &_
	"	,( SELECT COUNT(*) + 1 FROM [dbo].[SP_PROGRAM_APP] where [ProgramIdx] = @Idx AND [AreaIdx] = @AreaIdx ) AS [AppCode] " &_
	"	,A.[Kind] " &_
	"	,A.[Class] " &_
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
		.Parameters.Append .CreateParameter( "@Idx"     ,adInteger , adParamInput , 0 , programIdx )
		.Parameters.Append .CreateParameter( "@AreaIdx" ,adInteger , adParamInput , 0 , areaIdx )
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


Sub Inert()

	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"declare @ERR int ;" &_
	"set @ERR = 0 " &_

	"declare @State int,@ProgramIdx int,@AreaIdx int,@UserIdx int,@PayMode varchar(50),@Snumber varchar(50),@appCount varchar(3),@Snumber_last varchar(50) ;" &_
	"set @State      = ? " &_
	"set @ProgramIdx = ? " &_
	"set @AreaIdx    = ? " &_
	"set @UserIdx    = ? " &_
	"set @PayMode    = ? " &_
	"set @Snumber    = ? " &_
	"set @appCount   = ( SELECT COUNT(*) + 1 FROM [dbo].[SP_PROGRAM_APP] where [ProgramIdx] = @ProgramIdx AND [AreaIdx] = @AreaIdx ) " &_
	"set @Snumber_last = @Snumber + (REPLICATE('0', 3-LEN(@appCount)) + @appCount) " &_
	
	"IF @PayMode = 'SC0040' " &_
	"BEGIN " &_
	"set @Snumber_last = '' " &_
	"END " &_

	"BEGIN TRAN " &_

	"INSERT INTO [dbo].[SP_Pay_LGD]( " &_
	"	 [UserIdx] " &_
	"	,[LGD_RESPCODE] " &_
	"	,[LGD_RESPMSG] " &_
	"	,[LGD_PAYKEY] " &_
	"	,[LGD_MID] " &_
	"	,[LGD_OID] " &_
	"	,[LGD_AMOUNT] " &_
	"	,[LGD_TID] " &_
	"	,[LGD_PAYTYPE] " &_
	"	,[LGD_PAYDATE] " &_
	"	,[LGD_TIMESTAMP] " &_
	"	,[LGD_BUYER] " &_
	"	,[LGD_PRODUCTINFO] " &_
	"	,[LGD_BUYERID] " &_
	"	,[LGD_BUYERPHONE] " &_
	"	,[LGD_BUYEREMAIL] " &_
	"	,[LGD_BUYERSSN] " &_
	"	,[LGD_FINANCECODE] " &_
	"	,[LGD_FINANCENAME] " &_
	"	,[LGD_FINANCEAUTHNUM] " &_
	"	,[LGD_ESCROWYN] " &_
	"	,[LGD_CASHRECEIPTNUM] " &_
	"	,[LGD_CASHRECEIPTSELFYN] " &_
	"	,[LGD_CASHRECEIPTKIND] " &_
	"	,[LGD_CARDNUM] " &_
	"	,[LGD_CARDINSTALLMONTH] " &_
	"	,[LGD_CARDNOINTYN] " &_
	"	,[LGD_AFFILIATECODE] " &_
	"	,[LGD_CARDGUBUN1] " &_
	"	,[LGD_CARDGUBUN2] " &_
	"	,[LGD_CARDACQUIRER] " &_
	"	,[LGD_PCANCELFLAG] " &_
	"	,[LGD_PCANCELSTR] " &_
	"	,[LGD_TRANSAMOUNT] " &_
	"	,[LGD_EXCHANGERATE] " &_
	"	,[LGD_ACCOUNTNUM] " &_
	"	,[LGD_ACCOUNTOWNER] " &_
	"	,[LGD_PAYER] " &_
	"	,[LGD_CASTAMOUNT] " &_
	"	,[LGD_CASCAMOUNT] " &_
	"	,[LGD_CASFLAG] " &_
	"	,[LGD_CASSEQNO] " &_
	"	,[LGD_SAOWNER] " &_
	"	,[LGD_OCBAMOUNT] " &_
	"	,[LGD_OCBSAVEPOINT] " &_
	"	,[LGD_OCBTOTALPOINT] " &_
	"	,[LGD_OCBUSABLEPOINT] " &_
	"	,[LGD_OCBTID] " &_
	"	,[Indate] " &_
	")VALUES( " &_
	"	 @UserIdx " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,@PayMode " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,? " &_
	"	,getdate() " &_
	")" &_

	"INSERT INTO [dbo].[SP_PROGRAM_APP](  " &_
	"	 [State] " &_
	"	,[ProgramIdx] " &_
	"	,[AreaIdx] " &_
	"	,[UserIdx] " &_
	"	,[InData]" &_
	"	,[LgdIdx] " &_
	"	,[PayMode] " &_
	"	,[Snumber] " &_
	")VALUES( " &_
	"	 @State " &_
	"	,@ProgramIdx " &_
	"	,@AreaIdx " &_
	"	,@UserIdx " &_
	"	,getDate() " &_
	"	,SCOPE_IDENTITY() " &_
	"	,@PayMode " &_
	"	,@Snumber_last " &_
	") " &_

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
		.Parameters.Append .CreateParameter( "@State"     ,adInteger , adParamInput , 0 , State )
		.Parameters.Append .CreateParameter( "@ProgramIdx",adInteger , adParamInput , 0 , programIdx )
		.Parameters.Append .CreateParameter( "@AreaIdx"   ,adInteger , adParamInput , 0 , areaIdx )
		.Parameters.Append .CreateParameter( "@UserIdx"   ,adInteger , adParamInput , 0 , Session("UserIdx") )
		.Parameters.Append .CreateParameter( "@PayMode"   ,adVarChar , adParamInput ,50 , xpay.Response("LGD_PAYTYPE",0) )
		.Parameters.Append .CreateParameter( "@Snumber"   ,adVarChar , adParamInput ,50 , sNumber )

		.Parameters.Append .CreateParameter( "@LGD_RESPCODE"          ,adVarChar , adParamInput , 4 , xpay.Response("LGD_RESPCODE",0) )
		.Parameters.Append .CreateParameter( "@LGD_RESPMSG"           ,adVarChar , adParamInput , 512 , xpay.Response("LGD_RESPMSG",0) )
		.Parameters.Append .CreateParameter( "@LGD_PAYKEY"            ,adVarChar , adParamInput , 100 , LGD_PAYKEY )
		.Parameters.Append .CreateParameter( "@LGD_MID"               ,adVarChar , adParamInput , 15 , xpay.Response("LGD_MID",0) )
		.Parameters.Append .CreateParameter( "@LGD_OID"               ,adVarChar , adParamInput , 64 , xpay.Response("LGD_OID",0) )
		.Parameters.Append .CreateParameter( "@LGD_AMOUNT"            ,adVarChar , adParamInput , 12 , xpay.Response("LGD_AMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_TID"               ,adVarChar , adParamInput , 24 , xpay.Response("LGD_TID",0) )
		'.Parameters.Append .CreateParameter( "@LGD_PAYTYPE"           ,adVarChar , adParamInput , 6 , xpay.Response("LGD_PAYTYPE",0) )
		.Parameters.Append .CreateParameter( "@LGD_PAYDATE"           ,adVarChar , adParamInput , 14 , xpay.Response("LGD_PAYDATE",0) )
		.Parameters.Append .CreateParameter( "@LGD_TIMESTAMP"         ,adVarChar , adParamInput , 14 , xpay.Response("LGD_TIMESTAMP",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYER"             ,adVarChar , adParamInput , 10 , xpay.Response("LGD_BUYER",0) )
		.Parameters.Append .CreateParameter( "@LGD_PRODUCTINFO"       ,adVarChar , adParamInput , 128 , xpay.Response("LGD_PRODUCTINFO",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYERID"           ,adVarChar , adParamInput , 15 , xpay.Response("LGD_BUYERID",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYERPHONE"        ,adVarChar , adParamInput , 15 , xpay.Response("LGD_BUYERPHONE",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYEREMAIL"        ,adVarChar , adParamInput , 40 , xpay.Response("LGD_BUYEREMAIL",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYERSSN"          ,adVarChar , adParamInput , 13 , xpay.Response("LGD_BUYERSSN",0) )
		.Parameters.Append .CreateParameter( "@LGD_FINANCECODE"       ,adVarChar , adParamInput , 10 , xpay.Response("LGD_FINANCECODE",0) )
		.Parameters.Append .CreateParameter( "@LGD_FINANCENAME"       ,adVarChar , adParamInput , 20 , xpay.Response("LGD_FINANCENAME",0) )
		.Parameters.Append .CreateParameter( "@LGD_FINANCEAUTHNUM"    ,adVarChar , adParamInput , 20 , xpay.Response("LGD_FINANCEAUTHNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_ESCROWYN"          ,adVarChar , adParamInput , 1 , xpay.Response("LGD_ESCROWYN",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASHRECEIPTNUM"    ,adVarChar , adParamInput , 10 , xpay.Response("LGD_CASHRECEIPTNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASHRECEIPTSELFYN" ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CASHRECEIPTSELFYN",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASHRECEIPTKIND"   ,adVarChar , adParamInput , 4 , xpay.Response("LGD_CASHRECEIPTKIND",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDNUM"           ,adVarChar , adParamInput , 20 , xpay.Response("LGD_CARDNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDINSTALLMONTH"  ,adVarChar , adParamInput , 2 , xpay.Response("LGD_CARDINSTALLMONTH",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDNOINTYN"       ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CARDNOINTYN",0) )
		.Parameters.Append .CreateParameter( "@LGD_AFFILIATECODE"     ,adVarChar , adParamInput , 10 , xpay.Response("LGD_AFFILIATECODE",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDGUBUN1"        ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CARDGUBUN1",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDGUBUN2"        ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CARDGUBUN2",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDACQUIRER"      ,adVarChar , adParamInput , 2 , xpay.Response("LGD_CARDACQUIRER",0) )
		.Parameters.Append .CreateParameter( "@LGD_PCANCELFLAG"       ,adVarChar , adParamInput , 1 , xpay.Response("LGD_PCANCELFLAG",0) )
		.Parameters.Append .CreateParameter( "@LGD_PCANCELSTR"        ,adVarChar , adParamInput , 128 , xpay.Response("LGD_PCANCELSTR",0) )
		.Parameters.Append .CreateParameter( "@LGD_TRANSAMOUNT"       ,adVarChar , adParamInput , 12 , xpay.Response("LGD_TRANSAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_EXCHANGERATE"      ,adVarChar , adParamInput , 10 , xpay.Response("LGD_EXCHANGERATE",0) )
		.Parameters.Append .CreateParameter( "@LGD_ACCOUNTNUM"        ,adVarChar , adParamInput , 20 , xpay.Response("LGD_ACCOUNTNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_ACCOUNTOWNER"      ,adVarChar , adParamInput , 40 , xpay.Response("LGD_ACCOUNTOWNER",0) )
		.Parameters.Append .CreateParameter( "@LGD_PAYER"             ,adVarChar , adParamInput , 40 , xpay.Response("LGD_PAYER",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASTAMOUNT"        ,adVarChar , adParamInput , 12 , xpay.Response("LGD_CASTAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASCAMOUNT"        ,adVarChar , adParamInput , 12 , xpay.Response("LGD_CASCAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASFLAG"           ,adVarChar , adParamInput , 10 , xpay.Response("LGD_CASFLAG",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASSEQNO"          ,adVarChar , adParamInput , 3 , xpay.Response("LGD_CASSEQNO",0) )
		.Parameters.Append .CreateParameter( "@LGD_SAOWNER"           ,adVarChar , adParamInput , 10 , "" )
		.Parameters.Append .CreateParameter( "@LGD_OCBAMOUNT"         ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBSAVEPOINT"      ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBSAVEPOINT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBTOTALPOINT"     ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBTOTALPOINT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBUSABLEPOINT"    ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBUSABLEPOINT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBTID"            ,adVarChar , adParamInput , 24 , xpay.Response("LGD_OCBTID",0) )		

		set objRs = .Execute
	End with
	call cmdclose()

	' �ʵ� �ε����� ���� ����.
	CALL setFieldValue(objRs, "FI")
	objRs.close	: Set objRs = Nothing
end Sub



Sub Inert_Fail()

	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"INSERT INTO [dbo].[SP_Pay_LGD_FAIL]( [UserIdx] ,[LGD_RESPCODE] ,[LGD_RESPMSG] ,[LGD_PAYKEY] ,[LGD_MID] ,[LGD_OID] ,[LGD_AMOUNT] ,[LGD_TID] ,[LGD_PAYTYPE] ,[LGD_PAYDATE] ,[LGD_TIMESTAMP] ,[LGD_BUYER] ,[LGD_PRODUCTINFO] ,[LGD_BUYERID] ,[LGD_BUYERPHONE] ,[LGD_BUYEREMAIL] ,[LGD_BUYERSSN] ,[LGD_FINANCECODE] ,[LGD_FINANCENAME] ,[LGD_FINANCEAUTHNUM] ,[LGD_ESCROWYN] ,[LGD_CASHRECEIPTNUM] ,[LGD_CASHRECEIPTSELFYN] ,[LGD_CASHRECEIPTKIND] ,[LGD_CARDNUM] ,[LGD_CARDINSTALLMONTH] ,[LGD_CARDNOINTYN] ,[LGD_AFFILIATECODE] ,[LGD_CARDGUBUN1] ,[LGD_CARDGUBUN2] ,[LGD_CARDACQUIRER] ,[LGD_PCANCELFLAG] ,[LGD_PCANCELSTR] ,[LGD_TRANSAMOUNT] ,[LGD_EXCHANGERATE] ,[LGD_ACCOUNTNUM] ,[LGD_ACCOUNTOWNER] ,[LGD_PAYER] ,[LGD_CASTAMOUNT] ,[LGD_CASCAMOUNT] ,[LGD_CASFLAG] ,[LGD_CASSEQNO] ,[LGD_SAOWNER] ,[LGD_OCBAMOUNT] ,[LGD_OCBSAVEPOINT] ,[LGD_OCBTOTALPOINT] ,[LGD_OCBUSABLEPOINT] ,[LGD_OCBTID] ,[Indate] " &_
	
	")VALUES( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,getdate() )"


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@UserIdx"               ,adInteger , adParamInput , 0 , Session("UserIdx") )

		.Parameters.Append .CreateParameter( "@LGD_RESPCODE"          ,adVarChar , adParamInput , 4 , xpay.Response("LGD_RESPCODE",0) )
		.Parameters.Append .CreateParameter( "@LGD_RESPMSG"           ,adVarChar , adParamInput , 512 , xpay.Response("LGD_RESPMSG",0) )
		.Parameters.Append .CreateParameter( "@LGD_PAYKEY"            ,adVarChar , adParamInput , 100 , LGD_PAYKEY )
		.Parameters.Append .CreateParameter( "@LGD_MID"               ,adVarChar , adParamInput , 15 , xpay.Response("LGD_MID",0) )
		.Parameters.Append .CreateParameter( "@LGD_OID"               ,adVarChar , adParamInput , 64 , xpay.Response("LGD_OID",0) )
		.Parameters.Append .CreateParameter( "@LGD_AMOUNT"            ,adVarChar , adParamInput , 12 , xpay.Response("LGD_AMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_TID"               ,adVarChar , adParamInput , 24 , xpay.Response("LGD_TID",0) )
		.Parameters.Append .CreateParameter( "@LGD_PAYTYPE"           ,adVarChar , adParamInput , 6 , xpay.Response("LGD_PAYTYPE",0) )
		.Parameters.Append .CreateParameter( "@LGD_PAYDATE"           ,adVarChar , adParamInput , 14 , xpay.Response("LGD_PAYDATE",0) )
		.Parameters.Append .CreateParameter( "@LGD_TIMESTAMP"         ,adVarChar , adParamInput , 14 , xpay.Response("LGD_TIMESTAMP",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYER"             ,adVarChar , adParamInput , 10 , xpay.Response("LGD_BUYER",0) )
		.Parameters.Append .CreateParameter( "@LGD_PRODUCTINFO"       ,adVarChar , adParamInput , 128 , xpay.Response("LGD_PRODUCTINFO",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYERID"           ,adVarChar , adParamInput , 15 , xpay.Response("LGD_BUYERID",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYERPHONE"        ,adVarChar , adParamInput , 15 , xpay.Response("LGD_BUYERPHONE",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYEREMAIL"        ,adVarChar , adParamInput , 40 , xpay.Response("LGD_BUYEREMAIL",0) )
		.Parameters.Append .CreateParameter( "@LGD_BUYERSSN"          ,adVarChar , adParamInput , 13 , xpay.Response("LGD_BUYERSSN",0) )
		.Parameters.Append .CreateParameter( "@LGD_FINANCECODE"       ,adVarChar , adParamInput , 10 , xpay.Response("LGD_FINANCECODE",0) )
		.Parameters.Append .CreateParameter( "@LGD_FINANCENAME"       ,adVarChar , adParamInput , 20 , xpay.Response("LGD_FINANCENAME",0) )
		.Parameters.Append .CreateParameter( "@LGD_FINANCEAUTHNUM"    ,adVarChar , adParamInput , 20 , xpay.Response("LGD_FINANCEAUTHNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_ESCROWYN"          ,adVarChar , adParamInput , 1 , xpay.Response("LGD_ESCROWYN",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASHRECEIPTNUM"    ,adVarChar , adParamInput , 10 , xpay.Response("LGD_CASHRECEIPTNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASHRECEIPTSELFYN" ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CASHRECEIPTSELFYN",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASHRECEIPTKIND"   ,adVarChar , adParamInput , 4 , xpay.Response("LGD_CASHRECEIPTKIND",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDNUM"           ,adVarChar , adParamInput , 20 , xpay.Response("LGD_CARDNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDINSTALLMONTH"  ,adVarChar , adParamInput , 2 , xpay.Response("LGD_CARDINSTALLMONTH",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDNOINTYN"       ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CARDNOINTYN",0) )
		.Parameters.Append .CreateParameter( "@LGD_AFFILIATECODE"     ,adVarChar , adParamInput , 10 , xpay.Response("LGD_AFFILIATECODE",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDGUBUN1"        ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CARDGUBUN1",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDGUBUN2"        ,adVarChar , adParamInput , 1 , xpay.Response("LGD_CARDGUBUN2",0) )
		.Parameters.Append .CreateParameter( "@LGD_CARDACQUIRER"      ,adVarChar , adParamInput , 2 , xpay.Response("LGD_CARDACQUIRER",0) )
		.Parameters.Append .CreateParameter( "@LGD_PCANCELFLAG"       ,adVarChar , adParamInput , 1 , xpay.Response("LGD_PCANCELFLAG",0) )
		.Parameters.Append .CreateParameter( "@LGD_PCANCELSTR"        ,adVarChar , adParamInput , 128 , xpay.Response("LGD_PCANCELSTR",0) )
		.Parameters.Append .CreateParameter( "@LGD_TRANSAMOUNT"       ,adVarChar , adParamInput , 12 , xpay.Response("LGD_TRANSAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_EXCHANGERATE"      ,adVarChar , adParamInput , 10 , xpay.Response("LGD_EXCHANGERATE",0) )
		.Parameters.Append .CreateParameter( "@LGD_ACCOUNTNUM"        ,adVarChar , adParamInput , 20 , xpay.Response("LGD_ACCOUNTNUM",0) )
		.Parameters.Append .CreateParameter( "@LGD_ACCOUNTOWNER"      ,adVarChar , adParamInput , 40 , xpay.Response("LGD_ACCOUNTOWNER",0) )
		.Parameters.Append .CreateParameter( "@LGD_PAYER"             ,adVarChar , adParamInput , 40 , xpay.Response("LGD_PAYER",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASTAMOUNT"        ,adVarChar , adParamInput , 12 , xpay.Response("LGD_CASTAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASCAMOUNT"        ,adVarChar , adParamInput , 12 , xpay.Response("LGD_CASCAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASFLAG"           ,adVarChar , adParamInput , 10 , xpay.Response("LGD_CASFLAG",0) )
		.Parameters.Append .CreateParameter( "@LGD_CASSEQNO"          ,adVarChar , adParamInput , 3 , xpay.Response("LGD_CASSEQNO",0) )
		.Parameters.Append .CreateParameter( "@LGD_SAOWNER"           ,adVarChar , adParamInput , 10 , "" )
		.Parameters.Append .CreateParameter( "@LGD_OCBAMOUNT"         ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBAMOUNT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBSAVEPOINT"      ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBSAVEPOINT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBTOTALPOINT"     ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBTOTALPOINT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBUSABLEPOINT"    ,adVarChar , adParamInput , 12 , xpay.Response("LGD_OCBUSABLEPOINT",0) )
		.Parameters.Append .CreateParameter( "@LGD_OCBTID"            ,adVarChar , adParamInput , 24 , xpay.Response("LGD_OCBTID",0) )
		.Execute
	End with
	call cmdclose()
end Sub

 %>
<html>
<head>
<TITLE> SPWEB ����������  </TITLE>

<script type="text/javascript">
if("<%=AlertMsg%>" != "") alert("<%=AlertMsg%>");

window.opener.PayResult('<%=ResultCode%>','<%=xpay.Response("LGD_FINANCENAME",0)%>','<%=xpay.Response("LGD_ACCOUNTNUM",0)%>');
window.close()
</script>
</head>