<!-- #include virtual = "/Include/Config.asp" -->
<%
	Dim NO_LEN		: NO_LEN = Req("NO_LEN")
	Dim CD_RURL		: CD_RURL = Req("CD_RURL")
	Dim NM_MEMO		: NM_MEMO = Req("NM_MEMO")
	Dim DT_FROM		: DT_FROM = Replace(Replace(Req("DT_FROM"),"-",""),"/","")
	Dim DT_TO		: DT_TO = Replace(Replace(Req("DT_TO"),"-",""),"/","")
	Dim YN_LOG		: YN_LOG = Replace(Replace(Req("YN_LOG"),"-",""),"/","")


	Dim adoCmd

	SET adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB_InTheF
		.CommandType		= adCmdStoredProc
		.CommandText		= "USP_LINKURL_INS"
		.Parameters.Append	.CreateParameter("@NO_LEN"		, adInteger,	adParamInput,		0,		NO_LEN)
		.Parameters.Append	.CreateParameter("@CD_RURL"		, adVarchar,	adParamInput,		200,	CD_RURL)
		.Parameters.Append	.CreateParameter("@NM_MEMO"		, adVarchar,	adParamInput,		200,	NM_MEMO)
		.Parameters.Append	.CreateParameter("@DT_FROM"		, adVarchar,	adParamInput,		8,		DT_FROM)
		.Parameters.Append	.CreateParameter("@DT_TO"		, adVarchar,	adParamInput,		8,		DT_TO)
		.Parameters.Append	.CreateParameter("@YN_LOG"		, adVarchar,	adParamInput,		8,		YN_LOG)
		.Parameters.Append	.CreateParameter("@CD_USERID"	, adVarchar,	adParamInput,		50,		admin_PrimaryID)
		.Parameters.Append	.CreateParameter("@CD_URLKEY"	, adVarchar,	adParamInputOutput,	50,		"")
	
		.Execute

		end with
	SET adoCmd = Nothing

	Response.Redirect "linkurl.asp"
%>