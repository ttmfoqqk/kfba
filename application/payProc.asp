<!-- #include file = "../_lib/header.asp" -->
<!-- #include file = "../_lib/uploadUtil.asp" -->
<!-- #include file = "../_lib/pront.common.asp" -->
<%
Dim savePath : savePath = "\appMember/" '÷�� ������.
Set UPLOAD__FORM = Server.CreateObject("DEXT.FileUpload") 
UPLOAD__FORM.AutoMakeFolder = True 
UPLOAD__FORM.DefaultPath = UPLOAD_BASE_PATH & savePath
UPLOAD__FORM.MaxFileLen		= 20 * 1024 * 1024 '10�ް�


Dim actType        : actType        = UPLOAD__FORM("actType")
Dim programIdx     : programIdx     = UPLOAD__FORM("programIdx")
Dim areaIdx        : areaIdx        = UPLOAD__FORM("areaIdx")
Dim payMethod      : payMethod      = UPLOAD__FORM("MypayMethod")

Dim LastName       : LastName       = UPLOAD__FORM("LastName")
Dim FirstName      : FirstName      = UPLOAD__FORM("FirstName")

Dim PhotoName      : PhotoName      = UPLOAD__FORM("PhotoName")
Dim oldPhotoName   : oldPhotoName   = UPLOAD__FORM("oldPhotoName")

Dim UserIdx        : UserIdx        = UPLOAD__FORM("UserIdx")
Dim applicationKey : applicationKey = UPLOAD__FORM("applicationKey")

Dim GoUrl          : GoUrl           = "../mypage/"

Dim LGD_FINANCENAME : LGD_FINANCENAME = UPLOAD__FORM("LGD_FINANCENAME")
Dim LGD_ACCOUNTNUM  : LGD_ACCOUNTNUM  = UPLOAD__FORM("LGD_ACCOUNTNUM")


If Isnumeric( UserIdx ) And payMethod <> "" Then 
	'
Else
	Call msgbox("�߸��� ��� �Դϴ�.", true)
End If

Call Expires()
Call dbopen()

	if (alertMsg <> "")	then
		actType	= ""
	Elseif (actType = "INSERT") Then
		
		If payMethod = "SC0040" Then 
			'Call InsertPay()
			
			' �ߺ�
			'If FI_CntDuplicate > 0 Then 
			'	Call msgbox("�̹� ��ϵ� �������� �Դϴ�." & vbCrLf & vbCrLf & "�������������� Ȯ���� �ּ���.", true)
			'End If

			' ����
			'If FI_CK_EndDate < Left(Now(),10) Then
			'	Call msgbox("���� �����Ǿ����ϴ�. CK_EndDate : " & FI_CK_EndDate & " , Now : " & Left(Now(),10), true)
			'End If
			' ������
			'If FI_CK_StartDate > Left(Now(),10) Then 
			'	Call msgbox("���� �����Ⱓ�� �ƴմϴ�. CK_StartDate : " & FI_CK_StartDate & " , Now : " & Left(Now(),10), true)
			'End If
			' �ο�����
			'If int(FI_CK_MaxNumber) <= int(FI_CK_CNT_APP) Then 
			'	Call msgbox("���� �����ʰ�! MaxNumber : " & FI_CK_MaxNumber & " , CK_CNT_APP : " & FI_CK_CNT_APP , true)
			'End If

			GoUrl = "payBank.asp?applicationKey=" & applicationKey & "&LGD_FINANCENAME=" & LGD_FINANCENAME & "&LGD_ACCOUNTNUM=" & LGD_ACCOUNTNUM
		End If	
		
		'If PhotoName <>"" Then 
		'	If FILE_CHECK_EXT_JPG(PhotoName) = True Then
		'		If UPLOAD__FORM.MaxFileLen >= UPLOAD__FORM("PhotoName").FileLen Then 
		'			PhotoName = DextFileUpload("PhotoName",UPLOAD_BASE_PATH & savePath,0)
		'		Else
		'			Call msgbox("������ ũ��� 20MB �� �ѱ�� �����ϴ�.", true)
		'		End If
		'	Else
		'		Call msgbox("�߸��� �����Դϴ�. [asp,php,jsp,html,js] ������ ���ε� �Ҽ� �����ϴ�.", true)
		'	End If

		'	If oldPhotoName <> "" Then
		'		Set FSO = CreateObject("Scripting.FileSystemObject")
		'			If (FSO.FileExists(UPLOAD_BASE_PATH & savePath & oldPhotoName)) Then	' ���� �̸��� ������ ���� �� ����
		'				fso.deletefile(UPLOAD_BASE_PATH & savePath & oldPhotoName)
		'			End If
		'		set FSO = Nothing
		'	End If
		'Else
		'	PhotoName = oldPhotoName
		'End If

		
		
		'Call InsertPhoto()
		'alertMsg = "����ó�� �Ǿ����ϴ�."
		alertMsg = ""
	
	else
		alertMsg = "actType[" & actType & "]�� ���ǵ��� �ʾҽ��ϴ�."
	end If
	
Call dbclose()

'�Է�
Sub InsertPay()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("ADODB.Command")

	SQL = "SET NOCOUNT ON; " &_
	"DECLARE @ProgramIdx INT, @AreaIdx INT , @UserIdx INT ,@PayMode VARCHAR(50); " &_
	"SET @ProgramIdx = ?; " &_
	"SET @AreaIdx    = ?; " &_
	"SET @UserIdx    = ?; " &_
	"SET @PayMode    = ?; " &_
	
	"DECLARE @CntDuplicate INT , @CK_StartDate DATETIME , @CK_EndDate DATETIME , @CK_MaxNumber INT , @CK_CNT_APP INT ;" &_
	"SET @CntDuplicate = ( SELECT COUNT(*) FROM [dbo].[SP_PROGRAM_APP] WHERE [ProgramIdx] = @ProgramIdx AND [UserIdx] = @UserIdx AND [State] != 2 )  " &_
	
	"SELECT " &_
	"	 @CK_StartDate = A.[StartDate] " &_
	"	,@CK_EndDate   = A.[EndDate] " &_
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
	"AND A.[Idx] = @ProgramIdx " &_


	"IF @CntDuplicate = 0 AND @CK_StartDate <= CONVERT(varchar(10),GETDATE(),23) AND @CK_EndDate >= CONVERT(varchar(10),GETDATE(),23) AND @CK_MaxNumber > @CK_CNT_APP " &_
	"BEGIN "&_
	"	INSERT INTO [dbo].[SP_PROGRAM_APP]( " &_
	"		 [State] " &_
    "		,[ProgramIdx] " &_
    "		,[AreaIdx] " &_
    "		,[UserIdx] " &_
    "		,[InData] " &_
	"		,[PayMode]" &_
	"	)VALUES( " &_
	"		 1 " &_
    "		,@ProgramIdx " &_
    "		,@AreaIdx " &_
    "		,@UserIdx " &_
    "		,getDate() " &_
	"		,@PayMode " &_
	"	) " &_
	"END " &_

	"SELECT " &_
	"	 @CntDuplicate AS [CntDuplicate] " &_
	"	,@CK_StartDate AS [CK_StartDate] " &_
	"	,@CK_EndDate AS [CK_EndDate] " &_
	"	,@CK_MaxNumber AS [CK_MaxNumber] " &_
	"	,@CK_CNT_APP AS [CK_CNT_APP] "


	call cmdopen()
	with objCmd
		.CommandText = SQL
		.Parameters.Append .CreateParameter( "@ProgramIdx" ,adInteger , adParamInput, 0 , programIdx )
		.Parameters.Append .CreateParameter( "@AreaIdx"    ,adInteger , adParamInput, 0 , areaIdx )
		.Parameters.Append .CreateParameter( "@UserIdx"    ,adInteger , adParamInput, 0 , UserIdx )
		.Parameters.Append .CreateParameter( "@PayMode"    ,adVarChar , adParamInput,50 , payMethod )
		set objRs = .Execute
	End with
	call cmdclose()
	'���α׷�����
	CALL setFieldValue(objRs, "FI")
	Set objRs = Nothing
End Sub

'�Է�
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
		.Parameters.Append .CreateParameter( "@FirstName" ,adVarChar , adParamInput, 50 , FirstName )
		.Parameters.Append .CreateParameter( "@LastName"  ,adVarChar , adParamInput, 50 , LastName )
		.Parameters.Append .CreateParameter( "@Photo"     ,adVarChar , adParamInput, 200  , PhotoName )
		.Parameters.Append .CreateParameter( "@UserIdx"   ,adInteger , adParamInput, 0  , UserIdx )
		.Execute
	End with
	call cmdclose()
End Sub

%>

<!DOCTYPE html>
<HTML>
<HEAD>
	<META HTTP-EQUIV="Contents-Type" Contents="text/html; charset=euc-kr">
</head>
<script language=javascript>
	if ("<%=alertMsg%>" != "") alert('<%=alertMsg%>');
	top.location.href = "<%=GoUrl%>";
</script>
</html>