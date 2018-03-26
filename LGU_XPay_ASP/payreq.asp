<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<!-- #include file="./lgdacom/md5.asp" -->
<%
if session("UserIdx") = "" or IsNull(session("UserIdx")) Then
	With Response
	 .Write "<script type='text/javascript'>alert('로그인이 필요합니다.');window.opener.location.reload();window.close()</script>"
	 .End
	End With
end If


Dim savePath : savePath = "./appMember/" '첨부 저장경로.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath
UPLOAD__FORM.MaxFileLen		= 20 * 1024 * 1024 '10메가






'With Response
' .Write "<script type='text/javascript'>alert('현제 실결제가 아닌 서비스 테스트중 입니다.\n\n회원님들 께서는 \'임시 응시접수\' 로 이동하셔서 응시 부탁 드립니다.\n문의사항은 1800-6288번으로 문의주시기 바랍니다.\n\n불편을 드려 대단히 죄송합니다.');</script>"
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
	payMethodTxt = "카드결제"
ElseIf payMethod = "SC0030" Then 
	payMethodTxt = "실시간 계좌이체"
ElseIf payMethod = "SC0060" Then 
	payMethodTxt = "핸드폰결제"
ElseIf payMethod = "SC0040" Then 
	payMethodTxt = "가상계좌입금"
End If

If PhotoName <>"" Then 
	If FILE_CHECK_EXT_JPG(PhotoName) = True Then
'		If 0 = UPLOAD__FORM("PhotoName").FileLen Then 
'			With Response
'			 .Write "<script type='text/javascript'>alert('잘못된 파일입니다. [JPG,JPEG] 파일만 등록해주세요.');window.opener.location.reload();window.close()</script>"
'			 .End
'			End With
'		End If
		If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("PhotoName").FileLen Then 
			PhotoName = DextFileUpload("PhotoName",UPLOAD_BASE_PATH & savePath,0)
		Else
			With Response
			 .Write "<script type='text/javascript'>alert('파일의 크기는 20MB 를 넘길수 없습니다.');window.opener.location.reload();window.close()</script>"
			 .End
			End With
		End If
	Else

		'With Response
		 '.Write "<script type='text/javascript'>alert('잘못된 파일입니다. [JPG,JPEG,GIF,PNG] 파일만 등록해주세요.');window.opener.location.reload();window.close()</script>"
		'.End
		'End With

	End If
	If oldPhotoName <> "" Then
		Set FSO = CreateObject("Scripting.FileSystemObject")
			If (FSO.FileExists(UPLOAD_BASE_PATH & savePath & oldPhotoName)) Then	' 같은 이름의 파일이 있을 때 삭제
				fso.deletefile(UPLOAD_BASE_PATH & savePath & oldPhotoName)
			End If
		set FSO = Nothing
	End If
Else
	PhotoName = oldPhotoName
End If

'If PhotoName = "" Then 
'	With Response
'	 .Write "<script type='text/javascript'>alert('잘못된 파일입니다. [JPG,JPEG,GIF,PNG] 파일만 등록해주세요.');window.opener.location.reload();window.close()</script>"
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
	 .Write "<script type='text/javascript'>alert('잘못된 경로입니다.');window.close()</script>"
	 .End
	End With
end If

' 중복
If FI_CntDuplicate > 0 Then 
	With Response
	 .Write "<script type='text/javascript'>alert('이미 등록된 응시정보 입니다.\n\n마이페이지에서 확인해 주세요.');window.close()</script>"
	 .End
	End With
End If
' 같은종목 합격여부
'Response.write FI_CntDuplicate_program
If FI_CntDuplicate_program > 0 Then 
	With Response
	 .Write "<script type='text/javascript'>alert('이미 합격한 자격종목 입니다.\n\n마이페이지에서 확인해 주세요.');window.close()</script>"
	 .End
	End With
End If
' 마감
If FI_CK_EndDate < Left(Now(),10) Then
	With Response
	 .Write "<script type='text/javascript'>alert('응시 마감되었습니다.');window.close()</script>"
	 .End
	End With
End If
' 접수전
If FI_CK_StartDate > Left(Now(),10) Then 
	With Response
	 .Write "<script type='text/javascript'>alert('응시 접수기간이 아닙니다.');window.close()</script>"
	 .End
	End With
End If
' 인원제한
If int(FI_CK_MaxNumber) <= int(FI_CK_CNT_APP) Then 
	With Response
	 .Write "<script type='text/javascript'>alert('응시 정원초과!');window.close()</script>"
	 .End
	End With
End If

PrograName = FI_Name

If FI_Class = "1" Then
	PrograName = PrograName & " 1급"
ElseIf FI_Class = "2" Then
	PrograName = PrograName & " 2급"
ElseIf FI_Class = "3" Then
	PrograName = PrograName & " 3급"
End If

If FI_Kind = "1" Then
	PrograName = PrograName & " [필기]"
ElseIf FI_Kind = "2" Then
	PrograName = PrograName & " [실기]"
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
	'프로그램정보
	CALL setFieldValue(objRs, "FI")
	'회원정보
	set objRs = objRs.NextRecordset
	CALL setFieldValue(objRs, "USER")

	Set objRs = Nothing
End Sub

'/*
' * [결제 인증요청 페이지(STEP2-1)]
' *
' * 샘플페이지에서는 기본 파라미터만 예시되어 있으며, 별도로 필요하신 파라미터는 연동메뉴얼을 참고하시어 추가 하시기 바랍니다.
' */

'/*
' * 1. 기본결제 인증요청 정보 변경
' *
' * 기본정보를 변경하여 주시기 바랍니다.(파라미터 전달시 POST를 사용하세요)
' */

CST_PLATFORM               = "service"					'LG유플러스 결제 서비스 선택(test:테스트, service:서비스)
CST_MID                    = "soribiblue"            '상점아이디(LG유플러스으로 부터 발급받으신 상점아이디를 입력하세요)
																 '테스트 아이디는 't'를 반드시 제외하고 입력하세요.
if CST_PLATFORM = "test" then                                    '상점아이디(자동생성)
	LGD_MID = "t" & CST_MID
else
	LGD_MID = CST_MID
end If
oid = replace(date,"-","") & Hour(now) & Minute(now) & second(now) & Session("UserIdx")

LGD_OID                    = oid                                '주문번호(상점정의 유니크한 주문번호를 입력하세요)
LGD_AMOUNT                 = FI_Pay								'결제금액("," 를 제외한 결제금액을 입력하세요)
LGD_MERTKEY                = "15389449198e287da43d4785c51b9ca1" '[반드시 세팅]상점MertKey(mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다')
LGD_BUYER                  = trim(Left(USER_UserName,10))       '구매자명
LGD_PRODUCTINFO            = trim(PrograName)                   '상품명
LGD_BUYEREMAIL             = trim(USER_UserEmail)               '구매자 이메일
LGD_TIMESTAMP              = year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2) '타임스탬프
LGD_CUSTOM_SKIN            = "red"                               '상점정의 결제창 스킨 (red, blue, cyan, green, yellow)
LGD_BUYERID          	   = trim(Left(USER_UserId,15))          '구매자 아이디
LGD_BUYERIP          	   = trim(g_uip)      	                 '구매자IP
LGD_CUSTOM_USABLEPAY	   = payMethod                           '"SC0010-SC0030-SC0060"             '결제수단 
LGD_BUYERPHONE			   = Trim(USER_UserHphone1) &"-"&Trim(USER_UserHphone2) &"-"&Trim(USER_UserHphone3)             '전화번호
LGD_DISPLAY_BUYERPHONE	   = "Y"								 '휴대폰번호입력여부
LGD_DISPLAY_BUYEREMAIL	   = "Y"								 '이메일주소입력여부
LGD_AUTOFILLYN_BUYER	   = "Y"								 '구매자명 자동채움
LGD_CASHRECEIPTYN		   = "N"								 '현금영수증발급 사용여부
LGD_INSTALLRANGE           = "0"								 '할부 0:2:3:4:5:6:7:8:9:10:11:12

LGD_DISABLECARD            = ""
LGD_DISPLAY_ACCOUNTPID     = "N"


'/*
' * 가상계좌(무통장) 결제 연동을 하시는 경우 아래 LGD_CASNOTEURL 을 설정하여 주시기 바랍니다.
' */
LGD_CASNOTEURL             = "http://" & Request.ServerVariables("SERVER_NAME") & "/LGU_XPay_ASP/cas_noteurl.asp"

'/*
' *************************************************
' * 2. MD5 해쉬암호화 (수정하지 마세요) - BEGIN
' *
' * MD5 해쉬암호화는 거래 위변조를 막기위한 방법입니다.
' *************************************************
' *
' * 해쉬 암호화 적용( LGD_MID + LGD_OID + LGD_AMOUNT + LGD_TIMESTAMP + LGD_MERTKEY )
' * LGD_MID          : 상점아이디
' * LGD_OID          : 주문번호
' * LGD_AMOUNT       : 금액
' * LGD_TIMESTAMP    : 타임스탬프
' * LGD_MERTKEY      : 상점MertKey (mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다)
' *
' * MD5 해쉬데이터 암호화 검증을 위해
' * LG유플러스에서 발급한 상점키(MertKey)를 환경설정 파일(lgdacom/conf/mall.conf)에 반드시 입력하여 주시기 바랍니다.
' */
LGD_HASHDATA = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_TIMESTAMP & LGD_MERTKEY )
LGD_CUSTOM_PROCESSTYPE = "TWOTR"
'/*
' *************************************************
' * 2. MD5 해쉬암호화 (수정하지 마세요) - END
' *************************************************
' */
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=EUC-KR">
<TITLE> 한국외식음료협회 결제페이지  </TITLE>

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
 * 상점결제 인증요청후 PAYKEY를 받아서 최종결제 요청.
 */
function doPay_ActiveX(){
	document.getElementById('LGD_BUTTON2').innerHTML='결제 요청중 입니다.';
    ret = xpay_check(document.getElementById('LGD_PAYINFO'), '<%= CST_PLATFORM %>');

    if (ret=="00"){     //ActiveX 로딩 성공
        var LGD_RESPCODE        = dpop.getData('LGD_RESPCODE');       //결과코드
        var LGD_RESPMSG         = dpop.getData('LGD_RESPMSG');        //결과메세지

        if( "0000" == LGD_RESPCODE ) { //인증성공
            var LGD_PAYKEY      = dpop.getData('LGD_PAYKEY');         //LG유플러스 인증KEY
            var msg = "인증결과 : " + LGD_RESPMSG + "\n";
            msg += "LGD_PAYKEY : " + LGD_PAYKEY +"\n\n";
            document.getElementById('LGD_PAYKEY').value = LGD_PAYKEY;
            //alert(msg);
            document.getElementById('LGD_PAYINFO').submit();
        } else { //인증실패
            alert("인증이 실패하였습니다. " + LGD_RESPMSG);
            /*
             * 인증실패 화면 처리
             */
			 document.getElementById('LGD_BUTTON2').innerHTML='<img src="./img/btn.gif" onclick="doPay_ActiveX();" style="cursor:pointer;">';
        }
    } else {
        alert("LG U+ 전자결제를 위한 ActiveX Control이  설치되지 않았습니다.");
        /*
         * 인증실패 화면 처리
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
<div id="LGD_ACTIVEX_DIV"/> <!-- ActiveX 설치 안내 Layer 입니다. 수정하지 마세요. -->
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
								<td class=td_title><img src="./img/icon_arrow.gif"> &nbsp상품명</td>
								<td width=275><%=LGD_PRODUCTINFO%></td>
							</tr>
							<tr height="28">
								<td class=td_title><img src="./img/icon_arrow.gif"> &nbsp결제종류</td>
								<td><%=payMethodTxt%></td>
							</tr>
							<tr height="28">
								<td class=td_title><img src="./img/icon_arrow.gif"> &nbsp결제금액</td>
								<td class=td_cont_pink><%= FormatNumber(LGD_AMOUNT,0) %> (VAT포함)</td>
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
			<div id="LGD_BUTTON1">결제를 위한 모듈을 다운 중이거나, 모듈을 설치하지 않았습니다. </div>
			<div id="LGD_BUTTON2" style="display:none"><img src="./img/btn.gif" onclick="doPay_ActiveX();" style="cursor:pointer;"></div>
		</td>
	</tr>
	<tr>
		<td align=center style="padding:0px 0px 20px 0px;">※ 결제상의 오류가 있으면 TEL : <span class=td_cont_pink><b>1800-6288</b></span> 번으로 연락주시기 바랍니다.</td>
	</tr>
	<tr>
		<td height="18" bgcolor="617bff">&nbsp;</td>
	</tr>


</table>
<br>

<br>
<input type="hidden" name="CST_PLATFORM"                value="<%= CST_PLATFORM %>">                   <!-- 테스트, 서비스 구분 -->
<input type="hidden" name="CST_MID"                     value="<%= CST_MID %>">                        <!-- 상점아이디 -->
<input type="hidden" name="LGD_MID"                     value="<%= LGD_MID %>">                        <!-- 상점아이디 -->
<input type="hidden" name="LGD_OID"                     value="<%= LGD_OID %>">                        <!-- 주문번호 -->
<input type="hidden" name="LGD_BUYER"                   value="<%= LGD_BUYER %>">                      <!-- 구매자 -->
<input type="hidden" name="LGD_PRODUCTINFO"             value="<%= LGD_PRODUCTINFO %>">                <!-- 상품정보 -->
<input type="hidden" name="LGD_AMOUNT"                  value="<%= LGD_AMOUNT %>">                     <!-- 결제금액 -->
<input type="hidden" name="LGD_BUYEREMAIL"              value="<%= LGD_BUYEREMAIL %>">                 <!-- 구매자 이메일 -->
<input type="hidden" name="LGD_CUSTOM_SKIN"             value="<%= LGD_CUSTOM_SKIN %>">                <!-- 결제창 SKIN -->
<input type="hidden" name="LGD_CUSTOM_PROCESSTYPE"      value="<%= LGD_CUSTOM_PROCESSTYPE %>">         <!-- 트랜잭션 처리방식 -->
<input type="hidden" name="LGD_TIMESTAMP"               value="<%= LGD_TIMESTAMP %>">                  <!-- 타임스탬프 -->
<input type="hidden" name="LGD_HASHDATA"                value="<%= LGD_HASHDATA %>">                   <!-- MD5 해쉬암호값 -->
<input type="hidden" name="LGD_PAYKEY"                  id="LGD_PAYKEY">                               <!-- LG유플러스 PAYKEY(인증후 자동셋팅)-->
<input type="hidden" name="LGD_VERSION"         		value="ASP_XPay_1.0">						   <!-- 버전정보 (삭제하지 마세요) -->
<input type="hidden" name="LGD_BUYERIP"                 value="<%= LGD_BUYERIP %>">        			   <!-- 구매자IP -->
<input type="hidden" name="LGD_BUYERID"                 value="<%= LGD_BUYERID %>">           		   <!-- 구매자ID -->
<input type="hidden" name="LGD_CUSTOM_USABLEPAY"		value="<%= LGD_CUSTOM_USABLEPAY %>">           <!-- 상점정의결제가능수단 -->
<input type="hidden" name="LGD_BUYERSSN"				value="<%= LGD_BUYERSSN %>">				   <!-- 구매자주민번호 -->
<input type="hidden" name="LGD_CHECKSSNYN"				value="<%= LGD_CHECKSSNYN %>">				   <!-- 구매자주민번호 체크여부 -->

<input type="hidden" name="LGD_BUYERPHONE"				value="<%= LGD_BUYERPHONE %>">				   <!-- 구매자전화번호 -->
<input type="hidden" name="LGD_DISPLAY_BUYERPHONE"		value="<%= LGD_DISPLAY_BUYERPHONE %>">		   <!-- 휴대폰번호입력여부 -->
<input type="hidden" name="LGD_DISPLAY_BUYEREMAIL"		value="<%= LGD_DISPLAY_BUYEREMAIL %>">		   <!-- 이메일주소입력여부 -->
<input type="hidden" name="LGD_DISPLAY_ACCOUNTPID"		value="<%= LGD_DISPLAY_ACCOUNTPID %>">		   <!-- 가상계좌주민번호입력여부 -->
<input type="hidden" name="LGD_AUTOFILLYN_BUYER"		value="<%= LGD_AUTOFILLYN_BUYER %>">		   <!-- 구매자명 자동채움 -->
<input type="hidden" name="LGD_AUTOFILLYN_BUYERSSN"		value="<%= LGD_AUTOFILLYN_BUYERSSN %>">		   <!-- 구매자 주민번호 자동채움 -->
<input type="hidden" name="LGD_CASHRECEIPTYN"			value="<%= LGD_CASHRECEIPTYN %>">			   <!-- 현금영수증발급사용여부 -->
<input type="hidden" name="LGD_DISABLECARD"				value="<%= LGD_DISABLECARD %>">                <!-- 사용불가능카드 -->
<input type="hidden" name="LGD_INSTALLRANGE"			value="<%= LGD_INSTALLRANGE %>">                <!-- 표시할부개월수 -->

<input type="hidden" name="programIdx" value="<%= programIdx %>">     <!-- 프로그램 IDX -->
<input type="hidden" name="areaIdx"    value="<%= areaIdx %>">        <!-- 검정장 IDX -->

<!-- 가상계좌(무통장) 결제연동을 하시는 경우  할당/입금 결과를 통보받기 위해 반드시 LGD_CASNOTEURL 정보를 LG 유플러스에 전송해야 합니다 . -->
<input type="hidden" name="LGD_CASNOTEURL"           value="<%= LGD_CASNOTEURL %>">                 <!-- 가상계좌 NOTEURL -->


</form>
</body>
<!--  xpay.js는 반드시  body 밑에 두시기 바랍니다. -->
<!--  UTF-8 인코딩 사용 시는 xpay.js 대신 xpay_utf-8.js 을  호출하시기 바랍니다.-->
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