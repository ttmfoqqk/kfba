<!-- #include file = "../_lib/header.asp" -->
<!-- #include file="./lgdacom/md5.asp" -->
<%
Call Expires()
Call dbopen()
'/*
' * [상점 결제결과처리(DB) 페이지]
' *
' * 1) 위변조 방지를 위한 hashdata값 검증은 반드시 적용하셔야 합니다.
' *
' */
LGD_RESPCODE            = trim(request("LGD_RESPCODE"))             '// 응답코드: 0000(성공) 그외 실패
LGD_RESPMSG             = trim(request("LGD_RESPMSG"))              '// 응답메세지
LGD_MID                 = trim(request("LGD_MID"))                  '// 상점아이디
LGD_OID                 = trim(request("LGD_OID"))                  '// 주문번호
LGD_AMOUNT              = trim(request("LGD_AMOUNT"))               '// 거래금액
LGD_TID                 = trim(request("LGD_TID"))                  '// LG유플러스에서 부여한 거래번호
LGD_PAYTYPE             = trim(request("LGD_PAYTYPE"))              '// 결제수단코드
LGD_PAYDATE             = trim(request("LGD_PAYDATE"))              '// 거래일시(승인일시/이체일시)
LGD_HASHDATA            = trim(request("LGD_HASHDATA"))             '// 해쉬값
LGD_FINANCECODE         = trim(request("LGD_FINANCECODE"))          '// 결제기관코드(은행코드)
LGD_FINANCENAME         = trim(request("LGD_FINANCENAME"))          '// 결제기관이름(은행이름)
LGD_ESCROWYN            = trim(request("LGD_ESCROWYN"))             '// 에스크로 적용여부
LGD_TIMESTAMP           = trim(request("LGD_TIMESTAMP"))            '// 타임스탬프
LGD_ACCOUNTNUM          = trim(request("LGD_ACCOUNTNUM"))           '// 계좌번호(무통장입금)
LGD_CASTAMOUNT          = trim(request("LGD_CASTAMOUNT"))           '// 입금총액(무통장입금)
LGD_CASCAMOUNT          = trim(request("LGD_CASCAMOUNT"))           '// 현입금액(무통장입금)
LGD_CASFLAG             = trim(request("LGD_CASFLAG"))              '// 무통장입금 플래그(무통장입금) - 'R':계좌할당, 'I':입금, 'C':입금취소
LGD_CASSEQNO            = trim(request("LGD_CASSEQNO"))             '// 입금순서(무통장입금)
LGD_CASHRECEIPTNUM      = trim(request("LGD_CASHRECEIPTNUM"))       '// 현금영수증 승인번호
LGD_CASHRECEIPTSELFYN   = trim(request("LGD_CASHRECEIPTSELFYN"))    '// 현금영수증자진발급제유무 Y: 자진발급제 적용, 그외 : 미적용
LGD_CASHRECEIPTKIND     = trim(request("LGD_CASHRECEIPTKIND"))      '// 현금영수증 종류 0: 소득공제용 , 1: 지출증빙용
LGD_PAYER            	= trim(request("LGD_PAYER"))             	'// 입금자명

'/*
' * 구매정보
' */
LGD_BUYER               = trim(request("LGD_BUYER"))                '// 구매자
LGD_PRODUCTINFO         = trim(request("LGD_PRODUCTINFO"))          '// 상품명
LGD_BUYERID             = trim(request("LGD_BUYERID"))              '// 구매자 ID
LGD_BUYERADDRESS        = trim(request("LGD_BUYERADDRESS"))         '// 구매자 주소
LGD_BUYERPHONE          = trim(request("LGD_BUYERPHONE"))           '// 구매자 전화번호
LGD_BUYEREMAIL          = trim(request("LGD_BUYEREMAIL"))           '// 구매자 이메일
LGD_BUYERSSN            = trim(request("LGD_BUYERSSN"))             '// 구매자 주민번호
LGD_PRODUCTCODE         = trim(request("LGD_PRODUCTCODE"))          '// 상품코드
LGD_RECEIVER            = trim(request("LGD_RECEIVER"))             '// 수취인
LGD_RECEIVERPHONE       = trim(request("LGD_RECEIVERPHONE"))        '// 수취인 전화번호
LGD_DELIVERYINFO        = trim(request("LGD_DELIVERYINFO"))         '// 배송지

'/*
' * hashdata 검증을 위한 mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다.
' * LG유플러스에서 발급한 상점키로 반드시 변경해 주시기 바랍니다.
' */
LGD_MERTKEY = "15389449198e287da43d4785c51b9ca1"  '//mertkey
LGD_HASHDATA2 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )
'LGD_HASHDATA2 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )

'/*
' * 상점 처리결과 리턴메세지
' *
' * OK  : 상점 처리결과 성공
' * 그외 : 상점 처리결과 실패
' *
' * ※ 주의사항 : 성공시 'OK' 문자이외의 다른문자열이 포함되면 실패처리 되오니 주의하시기 바랍니다.
' */
resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 결과값을 입력해 주시기 바랍니다."

if LGD_HASHDATA2 = LGD_HASHDATA then
	'//해쉬값 검증이 성공이면
	if LGD_RESPCODE = "0000" then
		'//결제가 성공이면
		if LGD_CASFLAG = "R" then
			'/*
			' * 무통장 할당 성공 결과 상점 처리(DB) 부분
			' * 상점 결과 처리가 정상이면 "OK"
			' */
			resultMSG = "OK"
		elseif LGD_CASFLAG = "I" then
			'/*
			' * 무통장 입금 성공 결과 상점 처리(DB) 부분
			' * 상점 결과 처리가 정상이면 "OK"
			' */
			
			Call getView()
			
			
			Dim isDBOK  : isDBOK  = true
			Dim sNumber : sNumber = ""
			Dim State   : State   = 1

			'수검번호 생성
			'응시년도2자리 + 응시월2자리 + 검정장3자리 + 검정과목1자리 + 필기/실기1자리 + 급수 1자리 + 등록번호3자리
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

			'생성된 수검번호 13자리 체크
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
			' * 무통장 입금취소 성공 결과 상점 처리(DB) 부분
			' * 상점 결과 처리가 정상이면 "OK"
			' */
			Call getView()

			Call UpdateC()
			If RESULT_ERR = "0" Then 
				resultMSG = "OK"
			End If

			
		end if
	else
		'//결제가 실패이면
		'/*
		' * 거래실패 결과 상점 처리(DB) 부분
		' * 상점결과 처리가 정상이면 "OK"
		' */
		resultMSG = "OK"
	end if
else
	'//해쉬값이 검증이 실패이면
	'/*
	' * hashdata검증 실패 로그를 처리하시기 바랍니다.
	' */
	resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 해쉬값 검증이 실패하였습니다."
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

	' 필드 인덱스값 변수 생성.
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

	' 필드 인덱스값 변수 생성.
	CALL setFieldValue(objRs, "RESULT")
	objRs.close	: Set objRs = Nothing
end Sub

%>