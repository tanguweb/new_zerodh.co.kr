<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/board/BoardMast.asp
	'#  Create		: / 2010.09.07
	'#	Info		: BoardMast 저장
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	Dim MODE, CD_BOARDCD, NM_BOARDNM, YN_USE, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, NO_VIEWORDER, CD_BOARDTYPE, YN_DOWNLOAD
	Dim ary_FIELDNM(15), ary_FIELDINFO(15), ary_CD_FIELDTYPE(15), ary_CD_REQUIRED(15)
	Dim objADO, adoCmd, Result
	Dim iloop, jloop

	MODE			= ReqF("hid_mode")
	CD_BOARDCD		= ReqF("CD_BOARDCD")
	NM_BOARDNM		= ReqF("NM_BOARDNM_"&CD_BOARDCD)
	YN_USE			= ReqF("YN_USE_"&CD_BOARDCD)
	YN_CNTVIEW		= ReqF("YN_CNTVIEW_"&CD_BOARDCD)
	YN_ANSWER		= ReqF("YN_ANSWER_"&CD_BOARDCD)
	YN_REPLY		= ReqF("YN_REPLY_"&CD_BOARDCD)
	YN_XML			= ReqF("YN_XML_"&CD_BOARDCD)
	NO_VIEWORDER	= ReqF("NO_VIEWORDER_"&CD_BOARDCD)
	CD_BOARDTYPE	= ReqF("CD_BOARDTYPE_"&CD_BOARDCD)
	YN_DOWNLOAD		= ReqF("YN_DOWNLOAD_"&CD_BOARDCD)

	for iloop=0 to 14
		ary_FIELDNM(iloop)		= ReqF("NM_FIELDNM" & (iloop+1) & "_" & CD_BOARDCD)
		ary_FIELDINFO(iloop)	= ReqF("NM_FIELDINFO" & (iloop+1) & "_" & CD_BOARDCD)
		ary_CD_FIELDTYPE(iloop) = ReqF("CD_FIELDTYPE" & (iloop+1) & "_" & CD_BOARDCD)
		ary_CD_REQUIRED(iloop) = ReqF("CD_REQUIRED" & (iloop+1) & "_" & CD_BOARDCD)
	next 

'	response.write "MODE:" & MODE & "--<br>"
'	response.write "NM_BOARDNM:" & NM_BOARDNM & "--<br>"
'	response.write "YN_USE:" & YN_USE & "--<br>"
'	response.write "YN_CNTVIEW:" & YN_CNTVIEW & "--<br>"
'	response.write "YN_ANSWER:" & YN_ANSWER & "--<br>"
'	response.write "YN_REPLY:" & YN_REPLY & "--<br>"
'	response.write "YN_XML:" & YN_XML & "--<br>"
'	response.write "NO_VIEWORDER:" & NO_VIEWORDER & "--<br>"
'	response.write "ary_FIELDNM(0)" & ary_FIELDNM(0) & "--<br>"
'	response.write "ary_FIELDINFO(0)" & ary_FIELDINFO(0) & "--<br>"
'	response.write "ary_CD_FIELDTYPE(0)" & ary_CD_FIELDTYPE(0) & "--<br>"
'	response.write "ary_FIELDNM(1)" & ary_FIELDNM(1) & "--<br>"
'	response.write "ary_FIELDINFO(1)" & ary_FIELDINFO(1) & "--<br>"
'	response.write "ary_CD_FIELDTYPE(1)" & ary_CD_FIELDTYPE(1) & "--<br>"
'	response.write "ary_FIELDNM(2)" & ary_FIELDNM(2) & "--<br>"
'	response.write "ary_FIELDINFO(2)" & ary_FIELDINFO(2) & "--<br>"
'	response.write "ary_CD_FIELDTYPE(2)" & ary_CD_FIELDTYPE(2) & "--<br>"
'
'	response.end 

	SET adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDMAST_EDIT"
		.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	1,		MODE)
		.Parameters.Append	.CreateParameter("@CD_BOARDCD"		, adInteger,	adParamInput,	,		CD_BOARDCD)
		.Parameters.Append	.CreateParameter("@NM_BOARDNM"		, adVarchar,	adParamInput,	200,	getText(NM_BOARDNM))
		.Parameters.Append	.CreateParameter("@CD_BOARDTYPE"	, adVarchar,	adParamInput,	20,		CD_BOARDTYPE)
		.Parameters.Append	.CreateParameter("@YN_USE"			, adChar,		adParamInput,	1,		YN_USE)
		.Parameters.Append	.CreateParameter("@YN_CNTVIEW"		, adChar,		adParamInput,	1,		YN_CNTVIEW)
		.Parameters.Append	.CreateParameter("@YN_ANSWER"		, adchar,		adParamInput,	1,		YN_ANSWER)
		.Parameters.Append	.CreateParameter("@YN_REPLY"		, adchar,		adParamInput,	1,		YN_REPLY)
		.Parameters.Append	.CreateParameter("@YN_DOWNLOAD"		, adchar,		adParamInput,	1,		YN_DOWNLOAD)
		.Parameters.Append	.CreateParameter("@YN_XML"			, adchar,		adParamInput,	1,		YN_XML)
		.Parameters.Append	.CreateParameter("@NO_VIEWORDER"	, adInteger,	adParamInput,	,		NO_VIEWORDER)
for iloop=0 to 14
		.Parameters.Append	.CreateParameter("@NM_FIELDNM_" & (iloop+1)	, adVarchar,	adParamInput,	100,	getText(ary_FIELDNM(iloop)))
next
for iloop=0 to 14
		.Parameters.Append	.CreateParameter("@NM_FIELDINFO_" & (iloop+1)	, adVarchar,	adParamInput,	200,	getText(ary_FIELDINFO(iloop)))
next
for iloop=0 to 14
		.Parameters.Append	.CreateParameter("@CD_FIELDTYPE_" & (iloop+1)	, adVarchar,	adParamInput,	50,		ary_CD_FIELDTYPE(iloop))
Next
for iloop=0 to 14
		.Parameters.Append	.CreateParameter("@CD_REQUIRED_" & (iloop+1)	, adVarchar,	adParamInput,	50,		ary_CD_REQUIRED(iloop))
next
		.Parameters.Append	.CreateParameter("@CD_USER"			, adVarchar,	adParamInput,	15,		m_UserID)
		.Parameters.Append	.CreateParameter("@NO_IPADDR"		, adVarchar,	adParamInput,	20,		m_IPAddr)
		.Parameters.Append  .CreateParameter("@Result"	, adVarChar,	adParamoutput, 20,	"")
		.Execute
		Result = .Parameters("@Result")
		end with
	SET adoCmd = Nothing
	
	if Result = "SUCCESS" then 
		response.write "<script>alert('게시판이 저장되었습니다.');document.location.href='./BoardMast.asp'</script>"
		response.end
	else
		response.write "<script>alert('["&Result&"] 게시판이 저장되지 못하였습니다. 관리자에게 문의해 주세요.');history.back();</script>"
		response.end
	end if 
%>