<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_FileSave.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<!-- #Include virtual = "/function/fn_SendMail_Form.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /admin/board/BoardCounsel_Proc.asp
	'#  Create		: / 2011.12.16
	'#	Info		: 1:1문의 답변 등록/수정
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	

	Dim adoCmd, Result1, Result1_ID, Result1_DEPTH, Result1_SEQ, Result2, Rtn_string
	Dim MODE, CD_BOARDCD, CD_BOARDID, CD_BOARDKEY, NO_BOARD_DEPTH, NO_BOARD_SEQ, YN_USE, NO_CNTVIEW, YN_DISPLAY, NO_VIEWORDER, NO_DOWNLOAD, YN_XML, NM_XMLPATH, NM_TAG
	Dim CD_BOARDID_reply, NM_CONTENTS_reply
	Dim UploadForm, i, j
	Dim eMail, YN_EMAIL, emailTo

	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	UploadForm.DefaultPath = FileSaveRoot_Board

	Mode			= UploadForm("MODE")
	CD_BOARDCD		= UploadForm("CD_BOARDCD")
	CD_BOARDID		= UploadForm("CD_BOARDID")
	CD_BOARDKEY		= UploadForm("CD_BOARDKEY")
	NO_BOARD_DEPTH	= UploadForm("NO_BOARD_DEPTH")
	NO_BOARD_SEQ	= UploadForm("NO_BOARD_SEQ")
	YN_USE			= UploadForm("YN_USE")
	CD_BOARDID_reply	= UploadForm("CD_BOARDID_reply")
	NM_CONTENTS_reply	= UploadForm("NM_CONTENTS_reply")
	NM_TAG			= UploadForm("NM_TAG")
	NO_CNTVIEW		= UploadForm("NO_CNTVIEW")
	YN_DISPLAY		= UploadForm("YN_DISPLAY")
	NO_VIEWORDER	= UploadForm("NO_VIEWORDER")
	NO_DOWNLOAD		= UploadForm("NO_DOWNLOAD")
	YN_XML			= UploadForm("YN_XML")
	NM_XMLPATH		= UploadForm("NM_XMLPATH")
	
'	response.write Mode & "<br>"
'	response.write CD_BOARDCD & "<br>"
'	response.write CD_BOARDID & "<br>"
'	response.write CD_BOARDKEY & "<br>"
'	response.write NO_BOARD_DEPTH & "<br>"
'	response.write NO_BOARD_SEQ & "<br>"
'	response.write "CD_BOARDID_reply : " & CD_BOARDID_reply & "<br>"
'	response.write NM_CONTENTS_reply & "<br>"
'	response.write YN_DISPLAY & "<br>"
'	response.write NO_CNTVIEW & "<br>"
'	response.write NO_VIEWORDER & "<br>"
'	response.write NO_DOWNLOAD & "<br>"
'	response.write YN_USE & "<br>"
'	response.write NM_TAG & "<br>"
'	response.End 
	
	YN_EMAIL = UploadForm("YN_EMAIL")
	emailTo = UploadForm("emailTo")

	' 답변 저장/수정
	SET adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDDETAIL_COUNSEL_EDIT"
		.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	1, Mode)
		If Mode = "I" Then 
		.Parameters.Append	.CreateParameter("@CD_BOARDID"		, adInteger,	adParamInput,	 , CD_BOARDID)
		Else 
		.Parameters.Append	.CreateParameter("@CD_BOARDID"		, adInteger,	adParamInput,	 , getBoolean(CD_BOARDID_reply="",0,CD_BOARDID_reply))
		End If 
		.Parameters.Append	.CreateParameter("@CD_BOARDKEY"		, adInteger,	adParamInput,	 , CD_BOARDKEY)
		.Parameters.Append	.CreateParameter("@NO_BOARD_DEPTH"	, adInteger,	adParamInput,    , NO_BOARD_DEPTH)
		.Parameters.Append	.CreateParameter("@NO_BOARD_SEQ"	, adInteger,	adParamInput,    , NO_BOARD_SEQ)
		.Parameters.Append	.CreateParameter("@CD_BOARDCD"		, adInteger,	adParamInput,    , CD_BOARDCD)
		.Parameters.Append	.CreateParameter("@YN_DISPLAY"		, adchar,		adParamInput,	1, YN_DISPLAY)
		.Parameters.Append	.CreateParameter("@NO_VIEWCNT"		, adInteger,	adParamInput,    , NO_CNTVIEW)
		.Parameters.Append	.CreateParameter("@NO_VIEWORDER"	, adInteger,	adParamInput,	 , NO_VIEWORDER)
		.Parameters.Append	.CreateParameter("@NO_DOWNLOAD"		, adInteger,	adParamInput,	 , NO_DOWNLOAD)
		.Parameters.Append	.CreateParameter("@YN_USE"			, adchar,		adParamInput,	1, YN_USE)
		.Parameters.Append	.CreateParameter("@NM_TITLE"		, adVarchar,	adParamInput,	500, "1:1문의 답변입니다.")
		.Parameters.Append	.CreateParameter("@NM_CONTENTS"		, adVarchar,	adParamInput,	100000, NM_CONTENTS_reply)
		.Parameters.Append	.CreateParameter("@NM_TAG"			, adVarchar,	adParamInput,	1000, "")
for i=1 to 15
		.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, "")
next		
		.Parameters.Append	.CreateParameter("@CD_USER"			, adVarchar,	adParamInput,	15, m_UserID)
		.Parameters.Append	.CreateParameter("@NO_IPADDR"		, adVarchar,	adParamInput,	20, m_IPAddr)
		.Parameters.Append  .CreateParameter("@Result"			, adVarChar,	adParamoutput,  20	)
		.Parameters.Append  .CreateParameter("@Result_ID"		, adInteger,	adParamoutput  	)
		.Execute
		Result1			= .Parameters("@Result")
		Result1_ID		= .Parameters("@Result_ID")
		end with
	SET adoCmd = Nothing


	if Result1 = "SUCCESS" then 
		If MODE = "I" Then	' 등록
			Rtn_string = "<script>alert('답변글이 저장되었습니다.');document.location.href='BoardCounselView.asp?CD_BOARDID="&CD_BOARDID&"';</script>" 
		ElseIf MODE = "U" Then	' 수정
			Rtn_string = "<script>alert('답변글이 수정되었습니다.');document.location.href='BoardCounselView.asp?CD_BOARDID="&CD_BOARDID&"';</script>" 
		ElseIf MODE = "D" Then	' 삭제
			Rtn_string = "<script>alert('질문/답변 글이 삭제되었습니다.');document.location.href='BoardCounsel.asp?CD_BOARDCD="&CD_BOARDCD&"';</script>" 
		End If 

		'If YN_EMAIL = "Y" And emailTo <> "" Then 
		'	eMail = fn_SendMail_Form("공감헤어", "master@gonggamhair.com", emailTo, 0, 0, "[공감헤어] 안녕하세요. 공감헤어 입니다.^^", NM_CONTENTS_reply)	'// 답변 메일 발송
		'End If

		response.write Rtn_string 
		response.end 
	else
		response.write "<script>alert('정상적으로 등록되지 못하였습니다.');history.back();</script>"
		response.end 
	end if 
		
%>