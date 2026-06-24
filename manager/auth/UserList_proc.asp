<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%	
	'########################################################################################
	'#	File		: /auth/UserList_proc.asp
	'#  Create		: 조영준 / 2011.08.04
	'#	Info		: 관리자 등록/수정
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	Dim adoCmd, Result
	
	Dim MODE : MODE = ReqF("MODE")
	Dim CD_ADMINID
	Dim NM_ADMINPW
	Dim NM_NAME
	Dim NO_TEL
	Dim NM_EMAIL
	Dim YN_LOGIN
	Dim YN_USE

	If MODE = "I" Then 
		CD_ADMINID = ReqF("CD_ADMINID")
		NM_ADMINPW = ReqF("NM_ADMINPW")
		NM_NAME = ReqF("NM_NAME")
		NO_TEL = ReqF("NO_TEL")
		NM_EMAIL = ReqF("NM_EMAIL")
		YN_LOGIN = ReqF("YN_LOGIN")
		YN_USE = ReqF("YN_USE") 
	ElseIf MODE = "U" Then 
		CD_ADMINID = ReqF("update_CD_ADMINID")
		NM_ADMINPW = ReqF("NM_ADMINPW_" & CD_ADMINID)
		NM_NAME = ReqF("NM_NAME_" & CD_ADMINID)
		NO_TEL = ReqF("NO_TEL_" & CD_ADMINID)
		NM_EMAIL = ReqF("NM_EMAIL_" & CD_ADMINID)
		YN_LOGIN = ReqF("YN_LOGIN_" & CD_ADMINID)
		YN_USE = ReqF("YN_USE_" & CD_ADMINID) 

		
	End If 
'		Response.Write NM_ADMINPW
'		Response.End 
	If MODE = "I" Then 
		
		Set adoCmd	=	Server.CreateObject("ADODB.Command")
			with adoCmd
			.ActiveConnection	= m_DB
			.CommandType		= adCmdStoredProc
			.CommandText		= "SPA_ADMIN_EDIT"
			.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	1, MODE)
			.Parameters.Append	.CreateParameter("@CD_ADMINID"		, adVarChar,	adParamInput,	20, CD_ADMINID)
			.Parameters.Append	.CreateParameter("@NM_ADMINPW"		, adVarChar,	adParamInput,	20, NM_ADMINPW)
			.Parameters.Append	.CreateParameter("@NM_NAME"			, adVarChar,	adParamInput,   50, NM_NAME)
			.Parameters.Append	.CreateParameter("@NO_TEL"			, adVarChar,	adParamInput,   50, NO_TEL)
			.Parameters.Append	.CreateParameter("@NO_PHONE"		, adVarChar,	adParamInput,   20, "")
			.Parameters.Append	.CreateParameter("@NO_FAX"			, adVarChar,	adParamInput,   20, "")
			.Parameters.Append	.CreateParameter("@NM_EMAIL"		, adVarChar,	adParamInput,   50, NM_EMAIL)
			.Parameters.Append	.CreateParameter("@NM_URL"				, adVarChar,	adParamInput,   200, "")
			.Parameters.Append	.CreateParameter("@NM_PROFILE_IMAGE"	, adVarChar,	adParamInput,   200, "")
			.Parameters.Append	.CreateParameter("@YN_LOGIN"		, adchar,		adParamInput,	1, YN_LOGIN)
			.Parameters.Append	.CreateParameter("@YN_USE"			, adchar,		adParamInput,   1, YN_USE)
			.Parameters.Append	.CreateParameter("@CD_LEVEL"		, adInteger,	adParamInput,   , 0)
			.Parameters.Append  .CreateParameter("@Result"			, adVarChar,	adParamoutput,  20	)
			.Execute
			Result			= .Parameters("@Result")
			End with
		Set adoCmd = Nothing
		
		'Response.End 

		If Result = "SUCCESS" Then
			Response.Write "<script>alert('" & CD_ADMINID & "(" & NM_NAME & ") 님의 계정이 추가되었습니다.');parent.document.location.reload();</script>"
			Response.End 
		ElseIf Result = "900" Then 
			Response.Write "<script>															"
			Response.Write "	alert('동일한 ID가 존재합니다. \n\nID를 다시 확인해주세요.');	"
			Response.Write "	parent.document.getElementById('CD_ADMINID').select;			"
			Response.Write "	parent.document.location.relaod();								"
			Response.Write "</script>															"
			Response.End 
		ElseIf Result = "901" Then 
			Response.Write "<script>alert('[" & Result & "] 등록에 실패하였습니다. \n\n관리자에게 문의하세요.');parent.document.location.reload();</script>"
			Response.End 
		Else
			Response.Write "<script>alert('등록에 실패하였습니다. \n\n관리자에게 문의하세요.');parent.document.location.reload();</script>"
			Response.End 
		End If 

	ElseIf MODE = "U" Then 
		
		Set adoCmd	=	Server.CreateObject("ADODB.Command")
			with adoCmd
			.ActiveConnection	= m_DB
			.CommandType		= adCmdStoredProc
			.CommandText		= "SPA_ADMIN_EDIT"
			.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	1, MODE)
			.Parameters.Append	.CreateParameter("@CD_ADMINID"		, adVarChar,	adParamInput,	20, CD_ADMINID)
			.Parameters.Append	.CreateParameter("@NM_ADMINPW"		, adVarChar,	adParamInput,	20, NM_ADMINPW)
			.Parameters.Append	.CreateParameter("@NM_NAME"			, adVarChar,	adParamInput,   50, NM_NAME)
			.Parameters.Append	.CreateParameter("@NO_TEL"			, adVarChar,	adParamInput,   50, NO_TEL)
			.Parameters.Append	.CreateParameter("@NO_PHONE"		, adVarChar,	adParamInput,   20, "")
			.Parameters.Append	.CreateParameter("@NO_FAX"			, adVarChar,	adParamInput,   20, "")
			.Parameters.Append	.CreateParameter("@NM_EMAIL"		, adVarChar,	adParamInput,   50, NM_EMAIL)
			.Parameters.Append	.CreateParameter("@NM_URL"				, adVarChar,	adParamInput,   200, "")
			.Parameters.Append	.CreateParameter("@NM_PROFILE_IMAGE"	, adVarChar,	adParamInput,   200, "")
			.Parameters.Append	.CreateParameter("@YN_LOGIN"		, adchar,		adParamInput,	1, YN_LOGIN)
			.Parameters.Append	.CreateParameter("@YN_USE"			, adchar,		adParamInput,   1, YN_USE)
			.Parameters.Append	.CreateParameter("@CD_LEVEL"		, adInteger,	adParamInput,   , 0)
			.Parameters.Append  .CreateParameter("@Result"			, adVarChar,	adParamoutput,  20	)
			.Execute
			Result			= .Parameters("@Result")
			End with
		Set adoCmd = Nothing
		
		If Result = "SUCCESS" Then
			Response.Write "<script>alert('" & CD_ADMINID & "(" & NM_NAME & ") 님의 계정이 수정되었습니다.');parent.document.location.reload();</script>"
			Response.End 
		ElseIf Result = "902" Then 
			Response.Write "<script>alert('" & CD_ADMINID & "(" & NM_NAME & ") 님의 계정이 수정이 실패하였습니다. \n\n관리자에게 문의하세요.');parent.document.location.reload();</script>"
			Response.End 
		Else
			Response.Write "<script>alert('수정에 실패하였습니다. \n\n관리자에게 문의하세요.');parent.document.location.reload();</script>"
			Response.End 
		End If

	End If 
%>