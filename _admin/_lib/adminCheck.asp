<!-- #include file = "header.asp" -->
<%
'checkAdminLogin(g_host&g_url)

Call dbopen()
	'CALL getMemberdepth()	' ������ ����.
Call dbclose()

Sub getMemberdepth()
	SET objRs	= Server.CreateObject("ADODB.RecordSet")
	SET objCmd	= Server.CreateObject("adodb.command")
	with objCmd
		.ActiveConnection  = objConn
		.prepared          = true
		.CommandType       = adCmdStoredProc
		.CommandText       = "usp_admin_member_depth"

		.Parameters("@UID").value = Session("Admin_Id")

		Set objRs = .Execute
	End with
	set objCmd = nothing
	
	' �ʵ� �ε����� ���� ����.
	CALL setFieldValue(objRs, "MD")

	objRs.close	: Set objRs = Nothing
End Sub

'���� ���
Function ACheckpage(page,Plink)
	If page > 0 Then 
		If Plink = "MD_UL" Then	 'ȸ������
			If MD_U1 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/member/member_01_L.asp"
			ElseIf MD_U2 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/member/member_02_L.asp"
			ElseIf MD_U3 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/member/member_03_L.asp"
			ElseIf MD_U4 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/member/member_04_01_L.asp"
			Else
				ACheckpage = "javascript:alert('������ �����ϴ�')"
			End If

		ElseIf Plink = "MD_DL" Then	 '�ڷ��
			If MD_D1 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/library/list.asp?btype=LIBRARY.MED"
			ElseIf MD_D2 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/library/list.asp?btype=LIBRARY.PRC"
			ElseIf MD_D3 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/library/list.asp?btype=LIBRARY.BIO"
			ElseIf MD_D4 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/library/list.asp?btype=LIBRARY.EDU"
			ElseIf MD_D5 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/library/list.asp?btype=LIBRARY.LAW"
			ElseIf MD_D6 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/library/list.asp?btype=LIBRARY.ORG"
			Else
				ACheckpage = "javascript:alert('������ �����ϴ�')"
			End If
		ElseIf Plink = "MD_EL" Then	'���/�̺�Ʈ
			If MD_E1 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/event/event01.asp"
			ElseIf MD_E2 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/event/event04"
			ElseIf MD_E3 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/event/event02.asp?btype=EVENT.AFTER"
			Else
				ACheckpage = "javascript:alert('������ �����ϴ�')"
			End If
		ElseIf Plink = "MD_P1" Then	'�����а�
			If MD_E1 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/dept/dept_L.asp"
			Else
				ACheckpage = "javascript:alert('������ �����ϴ�')"
			End If
		ElseIf Plink = "MD_CL" Then	 'Ŀ�´�Ƽ
			If MD_C1 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/community/community01.asp?btype=NOTICE"
			ElseIf MD_C2 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/community/community02.asp?btype="
			ElseIf MD_C3 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/community/community03.asp?btype=NEWS"
			ElseIf MD_C4 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/community/community04.asp?btype=FREE"
			ElseIf MD_C5 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/community/community05.asp?btype=QNA"
			ElseIf MD_C6 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/community/community06.asp?btype="
			ElseIf MD_C7 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/community/community07.asp"
			Else
				ACheckpage = "javascript:alert('������ �����ϴ�')"
			End If
		ElseIf Plink = "MD_AL" Then	'������
			If MD_E1 > 0 Then 
				ACheckpage = C_ADMIN_FOLDER & "/admin/admin_L.asp"
			Else
				ACheckpage = "javascript:alert('������ �����ϴ�')"
			End If
		else
			ACheckpage = Plink
		End If
	Else
		ACheckpage = "javascript:alert('������ �����ϴ�')"
	End If
End Function

'��������
Function ACheckpageing(page)
	If page > 0 Then 
		response.write ""
	Else
		response.write "<script>alert('������ �����ϴ�.');history.go(-1)</script>"
	End If
End Function

If session("ADMIN_TODAY_PASS") <> "" Then 
	session("ADMIN_TODAY_PASS") = IIF(INSTR(g_url,"Eting_07_01")>0 Or INSTR(g_url,"Eting_07_success")>0 Or INSTR(g_url,"Eting_07_false")>0,session("ADMIN_TODAY_PASS"),"")
End If

response.redirect "../Admin/Admin_01_L.asp"
%>