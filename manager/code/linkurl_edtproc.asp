<!-- #include virtual = "/Include/Config.asp" -->
<%
	Dim CD_URLKEY	: CD_URLKEY = Req("CD_URLKEY")
	
	Dim fd

	fd = "NM_MEMO_" & CD_URLKEY
	Dim NM_MEMO		: NM_MEMO = Req(fd)

	fd = "CD_RURL_" & CD_URLKEY
	Dim CD_RURL		: CD_RURL = Req(fd)

	fd = "DT_FROM_" & CD_URLKEY
	Dim DT_FROM		: DT_FROM = Replace(Replace(Req(fd),"-",""),"/","")

	fd = "DT_TO_" & CD_URLKEY
	Dim DT_TO		: DT_TO = Replace(Replace(Req(fd),"-",""),"/","")

	fd = "YN_LOG_" & CD_URLKEY
	Dim YN_LOG		: YN_LOG = Replace(Replace(Req(fd),"-",""),"/","")

	fd = "YN_USE_" & CD_URLKEY
	Dim YN_USE		: YN_USE = Replace(Replace(Req(fd),"-",""),"/","")



	Dim adoCmd

	SET adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB_InTheF
		.CommandType		= adCmdStoredProc
		.CommandText		= "USP_LINKURL_UPD"
		.Parameters.Append	.CreateParameter("@CD_RURL"		, adVarchar,	adParamInput,		200,	CD_RURL)
		.Parameters.Append	.CreateParameter("@NM_MEMO"		, adVarchar,	adParamInput,		200,	NM_MEMO)
		.Parameters.Append	.CreateParameter("@DT_FROM"		, adVarchar,	adParamInput,		8,		DT_FROM)
		.Parameters.Append	.CreateParameter("@DT_TO"		, adVarchar,	adParamInput,		8,		DT_TO)
		.Parameters.Append	.CreateParameter("@YN_LOG"		, adVarchar,	adParamInput,		8,		YN_LOG)
		.Parameters.Append	.CreateParameter("@YN_USE"		, adVarchar,	adParamInput,		8,		YN_USE)
		.Parameters.Append	.CreateParameter("@CD_USERID"	, adVarchar,	adParamInput,		50,		admin_PrimaryID)
		.Parameters.Append	.CreateParameter("@CD_URLKEY"	, adVarchar,	adParamInput,		50,		CD_URLKEY)
	
		.Execute

		end with
	SET adoCmd = Nothing

	Response.Redirect "linkurl.asp"
%>