<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/board/BoardDetail_display_proc.asp
	'#  Create		: 조영준 / 2011.08.04
	'#	Info		: 게시글 리스트에서 전시여부 변경
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	Dim adoCmd, Result
	Dim CD_BOARDID : CD_BOARDID = ReqQ("CD_BOARDID")
	Dim YN_DISPLAY : YN_DISPLAY = ReqQ("YN_DISPLAY")

	Set adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDDETAIL_DISPLAY_EDIT"
		.Parameters.Append	.CreateParameter("@CD_BOARDID"		, adInteger,	adParamInput,	,	CD_BOARDID)
		.Parameters.Append	.CreateParameter("@YN_DISPLAY"		, adChar,		adParamInput,	1,	YN_DISPLAY)
		.Parameters.Append	.CreateParameter("@CD_USER"			, adVarchar,	adParamInput,	20,	m_UserID)
		.Parameters.Append	.CreateParameter("@NO_IPADDR"		, adVarchar,	adParamInput,	20,	m_IPAddr)
		.Parameters.Append  .CreateParameter("@Result"			, adVarChar,	adParamoutput,	20,	"")
		.Execute
		Result = .Parameters("@Result")
		End with
	Set adoCmd = Nothing
	
	If Result = "SUCCESS" Then 
		Response.Write "<script>alert('게시글이 수정되었습니다.');parent.document.location.reload();</script>"
		Response.End
	ElseIf Result = "900" Then 
		Response.Write "<script>alert('["&Result&"] 존재하지 않은 게시글 입니다.');history.back();</script>"
		Response.End
	ElseIf Result = "901" Then 
		Response.Write "<script>alert('["&Result&"] 게시글이 수정되지 못하였습니다. 관리자에게 문의해 주세요.');history.back();</script>"
		Response.End
	Else 
		Response.Write "<script>alert('게시글이 수정되지 못하였습니다. 관리자에게 문의해 주세요.');history.back();</script>"
		Response.End
	End If 
%>