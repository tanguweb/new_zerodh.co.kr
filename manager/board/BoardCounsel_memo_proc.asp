<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/board/Board_24_Proc.asp
	'#  Create		: 조영준 / 2011.08.16
	'#	Info		: 24시간 친절한 상담 등록/수정 
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	
	Dim adoCmd, Result
	
	Dim COUNSEL_GUBUN : COUNSEL_GUBUN = ReqF("COUNSEL_GUBUN")
	Dim CD_BOARDID	  : CD_BOARDID	  = ReqF("CD_BOARDID")
	Dim NM_MEMO		  : NM_MEMO		  = ReqF("NM_MEMO_" & CD_BOARDID)
	Dim NM_STATE	  : NM_STATE	  = ReqF("NM_STATE_" & CD_BOARDID)
	Dim ANSWER_POPUP  : ANSWER_POPUP  = ReqF("ANSWER_POPUP")	'답변하기 팝업에서 처리결과 변경시 ('Y') 를 받음

	'Response.Write CD_BOARDID & "<br />"
	'Response.Write NM_MEMO & "<br />"
	'Response.Write "<script>alert('" & CD_BOARDID & "\r\n" & NM_MEMO & "');</script>"

	Set adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDDETAIL_MEMO_EDIT"
		.Parameters.Append	.CreateParameter("@COUNSEL_GUBUN"	, adVarchar,	adParamInput,	20	,	COUNSEL_GUBUN)
		.Parameters.Append	.CreateParameter("@CD_BOARDID"		, adInteger,	adParamInput,		,	CD_BOARDID)
		.Parameters.Append	.CreateParameter("@NM_MEMO"			, adVarchar,	adParamInput,	1000,	getText(NM_MEMO))
		.Parameters.Append	.CreateParameter("@NM_STATE"		, adVarchar,	adParamInput,	1000,	NM_STATE)
		.Parameters.Append	.CreateParameter("@CD_USER"			, adVarchar,	adParamInput,	20	,	m_UserID)
		.Parameters.Append	.CreateParameter("@NO_IPADDR"		, adVarchar,	adParamInput,	20	,	m_IPAddr)
		.Parameters.Append  .CreateParameter("@Result"			, adVarChar,	adParamoutput,	20	,	"")
		.Execute
		Result = .Parameters("@Result")
		End with
	Set adoCmd = Nothing
	
	If Result = "SUCCESS" And COUNSEL_GUBUN = "memo" Then 
		Response.Write "<script type='text/javascript'>									"
		Response.Write "	alert('메모가 등록되었습니다.');							"
		Response.Write "	parent.document.getElementById('CD_BOARDID').value='';		"
		Response.Write "	parent.document.getElementById('COUNSEL_GUBUN').value='';	"
		If ANSWER_POPUP = "Y" Then
			'Response.Write "	parent.fn_parentReload();								"	'상담관리 페이지도 새로고침(상담관리 팝업창에 함수 있음 / _BoardCounsel_Popup.asp)
			Response.Write "	parent.window.close();								"	'상담관리 페이지닫기
		End If 
		Response.Write "	parent.document.location.reload();							"
		Response.Write "</script>														"
		Response.End
	ElseIf Result = "SUCCESS" And COUNSEL_GUBUN = "state" Then 
		Response.Write "<script type='text/javascript'>									"
		Response.Write "	alert('처리결과가 변경되었습니다.');						"
		Response.Write "	parent.document.getElementById('CD_BOARDID').value='';		"
		Response.Write "	parent.document.getElementById('COUNSEL_GUBUN').value='';	"
		If ANSWER_POPUP = "Y" Then
			'Response.Write "	parent.fn_parentReload();								"	'상담관리 페이지도 새로고침(상담관리 팝업창에 함수 있음 / _BoardCounsel_Popup.asp)
			Response.Write "	parent.window.close();								"	'상담관리 페이지닫기
		End If
		Response.Write "	parent.document.location.reload();							"
		Response.Write "</script>														"
		Response.End
	ElseIf Result = "900" Then 
		Response.Write "<script type='text/javascript'>									"
		Response.Write "	alert('["&Result&"] 존재하지 않은 게시글 입니다.');			"
		Response.Write "	parent.document.getElementById('CD_BOARDID').value='';		"
		Response.Write "	parent.document.getElementById('COUNSEL_GUBUN').value='';	"
		Response.Write "	history.back();												"
		Response.Write "</script>														"
		Response.End
	ElseIf Result = "901" Or Result = "902" Then 
		Response.Write "<script type='text/javascript'>															 "
		Response.Write "	alert('["&Result&"] 게시글이 수정되지 못하였습니다. \r\n관리자에게 문의해 주세요.'); "
		Response.Write "	parent.document.getElementById('CD_BOARDID').value='';								 "
		Response.Write "	parent.document.getElementById('COUNSEL_GUBUN').value='';							 "
		Response.Write "	history.back();																		 "
		Response.Write "</script>																				 "
		Response.End
	Else 
		Response.Write "<script type='text/javascript'>												"
		Response.Write "	alert('게시글이 수정되지 못하였습니다. \r\n관리자에게 문의해 주세요.');	"
		Response.Write "	parent.document.getElementById('CD_BOARDID').value='';					"
		Response.Write "	parent.document.getElementById('COUNSEL_GUBUN').value='';				"
		Response.Write "	history.back();															"
		Response.Write "</script>																	"
		Response.End
	End If 
%>	