<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/auth/adminmenu_Proc.assp
	'#  Create		: 조영준 / 2011.01.07
	'#	Info		: 프로그램 코드 관리
	'#	Update		:
	'#	Update Memo	:
	'########################################################################################

	'---------------------------------------------------
	' 변수선언 및 셋팅
	'---------------------------------------------------
	Dim MODE, RTN_STR, nm_string, adoCmd, Result
	Dim CD_FORM, NO_DEPTH, CD_PFORM, CD_PART, NM_FORM, NM_URL, CD_NAMESPACE, CD_EDITION, NO_ORDER, YN_USER, YN_USE, NM_ICON, CD_TARGET
	Dim iloop, jloop, kloop
	Dim objADO, Rs, SQL, SQL1, SQL2

	MODE		= ReqF("MODE")
	RTN_STR		= ReqF("RTN_STR")
	nm_string	= ReqF("nm_string")

	CD_FORM		= ReqF("CD_FORM_" & nm_string)
	NO_DEPTH	= ReqF("NO_DEPTH_" & nm_string)
	CD_PFORM	= ReqF("CD_PFORM_" & nm_string)
	CD_PART		= ReqF("CD_PART_" & nm_string)
	NM_FORM		= ReqF("NM_FORM_" & nm_string)
	NM_URL		= ReqF("NM_URL_" & nm_string)
	CD_NAMESPACE= ReqF("CD_NAMESPACE_" & nm_string)
	CD_EDITION	= ReqF("CD_EDITION_" & nm_string)
	'NM_ICON		= ReqF("NM_ICON_" & nm_string)
	NO_ORDER	= ReqF("NO_ORDER_" & nm_string)
	YN_USER		= ReqF("YN_USER_" & nm_string)
	YN_USE		= ReqF("YN_USE_" & nm_string)
	CD_TARGET	= ReqF("CD_TARGET_" & nm_string)

'	response.write "MODE		 : " & MODE			& "<br>"
'	response.write "RTN_STR		 : " & RTN_STR		& "<br>"
'	response.write "nm_string	 : " & nm_string	& "<br>"
'	response.write "CD_FORM		 : " & CD_FORM		& "<br>"
'	response.write "NO_DEPTH	 : " & NO_DEPTH		& "<br>"
'	response.write "CD_PFORM	 : " & CD_PFORM		& "<br>"
'	response.write "CD_PART		 : " & CD_PART		& "<br>"
'	response.write "NM_FORM		 : " & NM_FORM		& "<br>"
'	response.write "NM_URL		 : " & NM_URL		& "<br>"
'	response.write "CD_NAMESPACE : " & CD_NAMESPACE & "<br>"
'	response.write "CD_EDITION	 : " & CD_EDITION	& "<br>"
'	response.write "NO_ORDER	 : " & NO_ORDER		& "<br>"
'	response.write "YN_USER		 : " & YN_USER		& "<br>"
'	response.write "YN_USE		 : " & YN_USE		& "<br>"
'	response.End

	'// 최상위 메뉴의 상위 메뉴(CD_PFORM)는 자신의 메뉴코드 값
	If NO_DEPTH = "0" Then
		CD_PFORM = ""
	End If

	If MODE = "" Or RTN_STR = "" Or nm_string = "" Or CD_FORM = "" Or NO_DEPTH = "" Or NM_FORM = "" Or NO_ORDER = "" Or YN_USER = "" Or YN_USE = "" Then
		response.write "<script>alert('필요한 인수의 개수가 부족합니다.');parent.document.location.reload();</script>"
		response.End
	End If
	'---------------------------------------------------
	' 데이터
	'---------------------------------------------------

	Set objADO = new clsADO
	objADO.setConString(m_DB)

	'// 중복 체크
	If MODE = "I" Then
		SQL = ""
		SQL = SQL & " SELECT * FROM T_MENU WHERE CD_FORM = '" & CD_FORM & "' "
		objADO.setSQL(SQL)
		Set Rs = objADO.getRs()

		If Not Rs.EOF Then
			Response.Write "<script type='text/javascript'>alert('이미 등록되어진 메뉴 코드 입니다');parent.document.location.reload();</script>"
			Response.End
		End If
	End If


	'// SPA_TS_PG10_EDIT 호출 및 등록, 수정, 삭제
	SET adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_MENU_EDIT"
		.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	 1, MODE)
		.Parameters.Append	.CreateParameter("@CD_FORM"			, adVarchar,	adParamInput,	50, CD_FORM)
		.Parameters.Append	.CreateParameter("@CD_PFORM"		, adVarchar,	adParamInput,	50, CD_PFORM)
		.Parameters.Append	.CreateParameter("@NO_DEPTH"		, adInteger,	adParamInput,	  , NO_DEPTH)
		.Parameters.Append	.CreateParameter("@CD_PART"			, adVarchar,	adParamInput,   10, CD_PART)
		.Parameters.Append	.CreateParameter("@NM_FORM"			, adVarchar,	adParamInput,  100, NM_FORM)
		.Parameters.Append	.CreateParameter("@NM_URL"			, adVarchar,	adParamInput,  200, NM_URL)
		.Parameters.Append	.CreateParameter("@CD_NAMESPACE"	, adVarchar,	adParamInput,  100, CD_NAMESPACE)
		.Parameters.Append	.CreateParameter("@CD_EDITION"		, adVarchar,	adParamInput,	100, CD_EDITION)
		'.Parameters.Append	.CreateParameter("@NM_ICON"			, adVarchar,	adParamInput,	300, NM_ICON)
		.Parameters.Append	.CreateParameter("@NO_ORDER"		, adInteger,	adParamInput,	  , NO_ORDER)
		.Parameters.Append	.CreateParameter("@YN_USER"			, adChar,		adParamInput,	 1, YN_USER)
		.Parameters.Append	.CreateParameter("@YN_USE"			, adChar,		adParamInput,	 1, YN_USE)
		.Parameters.Append	.CreateParameter("@CD_TARGET"		, adVarchar,	adParamInput,	20, CD_TARGET)
		.Parameters.Append  .CreateParameter("@Result"			, adVarchar,	adParamoutput,   20	)
		.Execute
		Result				= .Parameters("@Result")
		End with
	SET adoCmd = Nothing

	'// 수정
	If Result = "SUCCESS" And MODE = "U" Then
		Response.Write "<script type='text/javascript'>alert('" & CD_FORM & "이(가) 수정 되었습니다.');parent.document.location.href='" & RTN_STR & "';</script>"
		Response.End
	'// 신규
	ElseIf Result = "SUCCESS" And MODE = "I" Then
		Response.Write "<script type='text/javascript'>alert('" & CD_FORM & "이(가) 등록 되었습니다.');parent.document.location.href='" & RTN_STR & "';</script>"
		Response.End
	'// 삭제
	ElseIf Result = "SUCCESS" And MODE = "D" Then
		Response.Write "<script type='text/javascript'>alert('" & CD_FORM & "이(가) 삭제 되었습니다.');parent.document.location.href='" & RTN_STR & "';</script>"
		Response.End
	End If

	'objADO.setSQL(SQL)
	'objADO.ExecuteQuery()

	Set objADO = Nothing

	'Response.Redirect RTN_STR
	'Response.End
%>
