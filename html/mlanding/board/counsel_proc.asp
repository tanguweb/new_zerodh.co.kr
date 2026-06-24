<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_FileSave.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp" -->
<!-- #include virtual = "/function/cafe24/sms/send.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		:
	'#  Create		: 2016.12
	'#	Info		:
	'#	Update		:
	'#	Update Memo	:
	'########################################################################################

	Dim adoCmd, Result1, Result1_ID, Result1_DEPTH, Result1_SEQ, Result2, Rtn_string
	Dim CD_BOARDKEY, NO_BOARD_DEPTH, NO_BOARD_SEQ, YN_USE, NO_CNTVIEW, YN_DISPLAY, NO_VIEWORDER, NO_DOWNLOAD, YN_XML, NM_XMLPATH
	Dim ary_FIELD(16), ary_FIELDTYPE(16)
	Dim UploadForm, arrFileNames(16), arrFileSize(16), arrFileType(16), arrFileTypechk, i, j, iCount, FilePath, FileNameOnly, FileExt, FileName

	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	UploadForm.DefaultPath = FileSaveRoot_Board
	UploadForm.MaxFileLen = 1024 * 1024 * 5	' 5MB

	Dim CD_BOARDCD : CD_BOARDCD = "1003"
	Dim MODE : MODE = UploadForm("MODE")
	If MODE = "" Then
		MODE = "I"
	End If

	Dim CD_BOARDID : CD_BOARDID = UploadForm("CD_BOARDID")
	If CD_BOARDID = "" Then
		CD_BOARDID = 0
	End If
	Dim NM_FIELD_1  : NM_FIELD_1 = getText(UploadForm("userName"))
	Dim NM_FIELD_2  : NM_FIELD_2 = getText(UploadForm("phone1")) & "-" & getText(UploadForm("phone2")) & "-" & getText(UploadForm("phone3"))
	Dim NM_FIELD_3  : NM_FIELD_3 = getText(UploadForm("email1")) & "@" & getText(UploadForm("email2"))
	'Dim NM_FIELD_4  : NM_FIELD_4 = getText(UploadForm("counselType"))
	Dim NM_FIELD_5  : NM_FIELD_5 = UploadForm("file1")
	Dim NM_FIELD_6  : NM_FIELD_6 = UploadForm("file2")
	Dim NM_FIELD_7  : NM_FIELD_7 = getText(UploadForm("password"))
	'Dim NM_FIELD_8  : NM_FIELD_8 = getText(UploadForm("sms_ok"))
	'Dim NM_FIELD_8  : NM_FIELD_8 = ""
	Dim NM_FIELD_9  : NM_FIELD_9 = getText(UploadForm("mail_ok"))
	Dim NM_FIELD_11 : NM_FIELD_11 = "READY"
	Dim NM_TITLE	: NM_TITLE = getText(UploadForm("title"))
	Dim NM_CONTENTS : NM_CONTENTS = Replace(getText(UploadForm("contents")) ,Chr(10),"<br />")
	Dim HTTP_REFERER : HTTP_REFERER = UploadForm("HTTP_REFERER")
	Dim HTTP_USER_AGENT : HTTP_USER_AGENT = UploadForm("HTTP_USER_AGENT")
	Dim HTTP_DEVICE_TYPE : HTTP_DEVICE_TYPE = UploadForm("HTTP_DEVICE_TYPE")
	Dim HTTP_MEDIA_TYPE : HTTP_MEDIA_TYPE = UploadForm("HTTP_MEDIA_TYPE")

'	If NM_FIELD_8 = "" Then
'		NM_FIELD_8 = "N"
'	End If

'	If NM_FIELD_9 = "" Then
'		NM_FIELD_9 = "N"
'	End If

	' 상담구분
	'Dim counselTypeName : counselTypeName = fn_CodeInfoWhere("COUNSEL_TYPE", " AND CD_CODE = '" & NM_FIELD_4 & "'")


'	Response.write "MODE : " & MODE & "<br />"
'	Response.write "NM_FIELD_1 : " & NM_FIELD_1 & "<br />"
'	Response.write "NM_FIELD_2 : " & NM_FIELD_2 & "<br />"
'	Response.write "NM_FIELD_3 : " & NM_FIELD_3 & "<br />"
'	Response.write "NM_FIELD_4 : " & NM_FIELD_4 & "<br />"
'	Response.write "NM_FIELD_7 : " & NM_FIELD_7 & "<br />"
'	Response.write "NM_FIELD_8 : " & NM_FIELD_8 & "<br />"
'	Response.write "NM_FIELD_9 : " & NM_FIELD_9 & "<br />"
'	Response.write "NM_TITLE : " & NM_TITLE & "<br />"
'	Response.write "NM_CONTENTS : " & NM_CONTENTS & "<br />"

	If MODE = "D" Then
		CD_BOARDKEY = CD_BOARDID
		YN_DISPLAY = "N"
		YN_USE = "N"
	Else
		'If MODE = "" Or NM_FIELD_1 = "" Or NM_FIELD_2 = "" Or NM_FIELD_3 = "" Or NM_FIELD_4 = "" Or NM_FIELD_7 = "" Or NM_TITLE = "" Or NM_CONTENTS = "" Then
		If MODE = "" Or NM_FIELD_1 = "" Or NM_FIELD_2 = "" Or NM_FIELD_3 = "" Or NM_FIELD_7 = "" Or NM_TITLE = "" Or NM_CONTENTS = "" Then
		'If MODE = "" Or NM_FIELD_1 = "" Or NM_FIELD_2 = "" Or NM_FIELD_7 = "" Or NM_TITLE = "" Or NM_CONTENTS = "" Then
			Response.Write "<script>alert('잘못된 접근 입니다.');document.location.href='counsel.html';</script>"
			Response.End
		End If

		CD_BOARDKEY = 0
		YN_DISPLAY = "Y"
		YN_USE = "Y"

	End If

	For i = 1 To 15
		ary_FIELD(i)		= UploadForm("NM_FIELD_" & i)
	Next

	ary_FIELD(1) = NM_FIELD_1
	ary_FIELD(2) = NM_FIELD_2
	ary_FIELD(3) = NM_FIELD_3
	'ary_FIELD(4) = NM_FIELD_4
	ary_FIELD(5) = NM_FIELD_5
	ary_FIELD(6) = NM_FIELD_6
	ary_FIELD(7) = NM_FIELD_7
'	ary_FIELD(8) = NM_FIELD_8
	ary_FIELD(9) = NM_FIELD_9
	ary_FIELD(11) = NM_FIELD_11


	' begin : 업로드 파일(file1) 유효성 체크 추가
	If UploadForm("file1") <> "" Then

		If UploadForm("file1").Count <> 1 Then
			Response.Write "<script type='text/javascript'>alert('잘못된 접근입니다.');history.back();</script>"
			Response.End 
		End If 

		Dim file1_filename : file1_filename = UploadForm("file1").FileName
		Dim file1_mimetype : file1_mimetype = UploadForm("file1").MimeType
		'Response.Write "file1_mimetype: " & file1_mimetype & "<br />"
		
		If checkFilename(file1_filename) = False Then
			Response.Write "<script type='text/javascript'>alert('업로드 불가능한 파일 이름입니다.');history.back();</script>"
			Response.End 
		End If 

		If file1_mimetype <> "image/jpeg" And file1_mimetype <> "image/png" And file1_mimetype <> "image/gif" Then 
			Response.Write "<script type='text/javascript'>alert('업로드 불가능한 파일 형식입니다.');history.back();</script>"
			Response.End 
		End If
	End If 
	' end : 업로드 파일 유효성(file1) 체크 추가

	If UploadForm("file1").FileLen > UploadForm.MaxFileLen Then
		Response.Write "<script type='text/javascript'>alert('첨부파일은 최대 5MB 까지 업로드 가능합니다.');history.back();</script>"
		Response.End 
	End If 
	
	' begin : 업로드 파일(file2) 유효성 체크 추가
	If UploadForm("file2") <> "" Then

		If UploadForm("file2").Count <> 1 Then
			Response.Write "<script type='text/javascript'>alert('잘못된 접근입니다.');history.back();</script>"
			Response.End 
		End If 

		Dim file2_filename : file2_filename = UploadForm("file2").FileName
		Dim file2_mimetype : file2_mimetype = UploadForm("file2").MimeType
		'Response.Write "file2_mimetype " & file2_mimetype & "<br />"
		
		If checkFilename(file2_filename) = False Then
			Response.Write "<script type='text/javascript'>alert('업로드 불가능한 파일 이름입니다.');history.back();</script>"
			Response.End 
		End If 

		If file2_mimetype <> "image/jpeg" And file2_mimetype <> "image/png" And file2_mimetype <> "image/gif" Then 
			Response.Write "<script type='text/javascript'>alert('업로드 불가능한 파일 형식입니다.');history.back();</script>"
			Response.End 
		End If
	End If 
	' end : 업로드 파일 유효성(file2) 체크 추가

	If UploadForm("file2").FileLen > UploadForm.MaxFileLen Then
		Response.Write "<script type='text/javascript'>alert('첨부파일은 최대 5MB 까지 업로드 가능합니다.');history.back();</script>"
		Response.End 
	End If 

	ary_FIELDTYPE(5) = "FILE"

	'Response.write "ary_FIELD(5) : [" & ary_FIELD(5) & "]<br />"
	'Response.write "hd_file1 : [" & UploadForm("hd_file1_o") & "]<br />"

	'Response.write "ary_FIELD(6) : [" & ary_FIELD(6) & "]<br />"
	'Response.write "hd_file2 : [" & UploadForm("hd_file2_o") & "]<br />"

	'Response.End

	If ary_FIELD(5) <> "" Then
		arrFileNames(5) = fn_FileSave (UploadForm, "file1", getBoolean(YN_XML="Y" and NM_XMLPATH<>"",FileSaveRoot & NM_XMLPATH,FileSaveRoot_Board))
	Else
		arrFileNames(5) = UploadForm("hd_file1_o")
	End If

	ary_FIELDTYPE(6) = "FILE"

	If ary_FIELD(6) <> "" Then
		arrFileNames(6) = fn_FileSave (UploadForm, "file2", getBoolean(YN_XML="Y" and NM_XMLPATH<>"",FileSaveRoot & NM_XMLPATH,FileSaveRoot_Board))
	Else
		arrFileNames(6) = UploadForm("hd_file2_o")
	End If

	' 게시글 저장
	Set adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDDETAIL_EDIT"
		.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	1, MODE)
		.Parameters.Append	.CreateParameter("@CD_BOARDID"		, adInteger,	adParamInput,	 , CD_BOARDID)
		.Parameters.Append	.CreateParameter("@CD_BOARDKEY"		, adInteger,	adParamInput,	 , CD_BOARDKEY)
		.Parameters.Append	.CreateParameter("@NO_BOARD_DEPTH"	, adInteger,	adParamInput,    , 1)
		.Parameters.Append	.CreateParameter("@NO_BOARD_SEQ"	, adInteger,	adParamInput,    , 1)
		.Parameters.Append	.CreateParameter("@CD_BOARDCD"		, adInteger,	adParamInput,    , CD_BOARDCD)
		.Parameters.Append	.CreateParameter("@YN_DISPLAY"		, adchar,		adParamInput,	1, YN_DISPLAY)
		.Parameters.Append	.CreateParameter("@NO_VIEWCNT"		, adInteger,	adParamInput,    , 0)
		.Parameters.Append	.CreateParameter("@NO_VIEWORDER"	, adInteger,	adParamInput,	 , 100)
		.Parameters.Append	.CreateParameter("@NO_DOWNLOAD"		, adInteger,	adParamInput,	 , 0)
		.Parameters.Append	.CreateParameter("@YN_USE"			, adchar,		adParamInput,	1, YN_USE)
		.Parameters.Append	.CreateParameter("@NM_TITLE"		, adVarchar,	adParamInput,	500, NM_TITLE)
		.Parameters.Append	.CreateParameter("@NM_CONTENTS"		, adVarchar,	adParamInput,	100000, NM_CONTENTS)
	For i = 1 To 15
		If ary_FIELDTYPE(i) = "FILE" Then
			.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, arrFileNames(i))
		Else
			.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, ary_FIELD(i))
		End If
	Next

		.Parameters.Append	.CreateParameter("@CD_USER"			, adVarchar,	adParamInput,	15, getBoolean(m_ServiceUserID<>"",m_ServiceUserID,"guest"))
		.Parameters.Append	.CreateParameter("@NO_IPADDR"		, adVarchar,	adParamInput,	20, m_IPAddr)
		.Parameters.Append	.CreateParameter("@HTTP_REFERER"		, adVarchar,	adParamInput,	5000, HTTP_REFERER)
		.Parameters.Append	.CreateParameter("@HTTP_USER_AGENT"		, adVarchar,	adParamInput,	5000, HTTP_USER_AGENT)
		.Parameters.Append	.CreateParameter("@HTTP_DEVICE_TYPE"	, adVarchar,	adParamInput,	200, HTTP_DEVICE_TYPE)
		.Parameters.Append	.CreateParameter("@HTTP_MEDIA_TYPE"		, adVarchar,	adParamInput,	200, HTTP_MEDIA_TYPE)
		.Parameters.Append  .CreateParameter("@Result"			, adVarChar,	adParamoutput,  20	)
		.Parameters.Append  .CreateParameter("@Result_ID"		, adInteger,	adParamoutput  	)
		.Execute
		Result1			= .Parameters("@Result")
		Result1_ID		= .Parameters("@Result_ID")
		End with
	Set adoCmd = Nothing

	If Result1 = "SUCCESS" Then

		Dim isSendMail : isSendMail = m_client_counsel_online_send_email 	'메일 발송할 것인지 여부

		If isSendMail = True Then
			Dim SendMail : SendMail = False
			Dim sFrom, sTo, objMail, MailConfig
			'sFrom = ""&NM_NAME&"" & "<webmaster@domain>" ' 보내는사람
			sFrom = "<"&CONFIG_SMTP_SENDUSERNAME&">" ' 보내는사람
			'sTo = "<master@peters.co.kr>"	' 받는사람



			Dim aryRecvEMailList, email_i
			aryRecvEMailList = fn_CodeInfoWhere("COUNSEL_MAIL_RECV", " AND NM_CODE1 = '"&CD_BOARDCD&"'")

			If IsArray(aryRecvEMailList) Then
				For email_i = 0 To UBound(aryRecvEMailList, 2)
					If sTo = "" Then
						sTo = sTo & "<" & aryRecvEMailList(5, email_i) & ">"
					Else
						sTo = sTo & ",<" & aryRecvEMailList(5, email_i) & ">"
					End If
				Next
			End If

			If sTo <> "" Then
				Set objMail = Server.CreateObject("CDO.Message") ' CDO 2.0(메일 보내기 컴포넌트 개체 생성)

				Set MailConfig = objMail.Configuration
				With MailConfig.Fields
				.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = CONFIG_SMTP_SENDUSING	' 1 : 로컬 SMTP) / 2 : 외부 SMTP)
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") =  CONFIG_SMTP_SMTPSERVERPICKUPDIRECTORY	' Pickup 디렉토리 설정
				'.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"  ' 호스트설정
				.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CONFIG_SMTP_SMTPSERVER
				.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CONFIG_SMTP_SMTPSERVERPORT	' SMTP Port
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" ) = CONFIG_SMTP_SMTPAUTHENTICATE
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusername" ) = CONFIG_SMTP_SENDUSERNAME
				.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = CONFIG_SMTP_SENDPASSWORD	'(리얼)
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = CONFIG_SMTP_SMTPCONNECTIONTIMEOUT	' 연결시간
				.Update
				End With
				Set MailConfig = Nothing

				objMail.From = sFrom ' 메일을 보내는 사람의 이메일 주소
				objMail.To = sTo ' 메일을 받는 사람의 이메일주소(여러사람일 경우는 ; 표시로 구분)
				'objMail.Subject = NM_FIELD_1 & "상담" ' 메일 제목
				objMail.Subject = "["&m_client_counsel_regist_email_title&"] 온라인상담 신청글이 등록되었습니다" & "(" & NM_FIELD_1 & ")"

				Dim mailBody

				mailBody = mailBody & m_client_counsel_regist_email_title & "에 온라인상담 신청글이 등록되었습니다.<br />"
				mailBody = mailBody & "신청글의 자세한 내용은 관리자 페이지에서 확인해주세요.<br /><br /><br />"
				mailBody = mailBody & "<strong>* 이름 : </strong>" & NM_FIELD_1 & "<br /><br />"
				mailBody = mailBody & "<strong>* 휴대폰번호 : </strong>" & NM_FIELD_2 & "<br /><br />"
				mailBody = mailBody & "<strong>* 이메일 : </strong>" & NM_FIELD_3 & "<br /><br />"
				'mailBody = mailBody & "<strong>* 상담과목 : </strong>" & counselTypeName(2, 0) & "<br /><br />"
				mailBody = mailBody & "<strong>* 제목 : </strong>" & NM_TITLE & "<br /><br />"
				mailBody = mailBody & "<strong>* 내용 : </strong>" & NM_CONTENTS & "<br /><br />"

				objMail.HTMLBody = mailBody

				'objMail.BodyPart.Charset="utf-8"
				'objMail.HTMLBodyPart.Charset="utf-8"

				objMail.send

				Set objMail = Nothing
				If Not Err Then
					SendMail = True
				Else
					SendMail = False
				End If
			End If

		End If


		Dim isSendSMS : isSendSMS = m_client_counsel_online_send_sms	'SMS 발송할 것인지 여부

		If MODE = "I" Then

			If isSendSMS = True Then

				Dim aryRecvSMSList, sms_i
				aryRecvSMSList = fn_CodeInfoWhere("COUNSEL_SMS_RECEIVER", " AND NM_CODE1 = '"&CD_BOARDCD&"'")

				'Response.Write "IsArray(aryRecvSMSList) : [" & IsArray(aryRecvSMSList) & "]<br />"

				If IsArray(aryRecvSMSList) Then
					For sms_i = 0 To UBound(aryRecvSMSList, 2)
						'Response.Write "aryRecvSMSList(5, sms_i) : [" & aryRecvSMSList(5, sms_i) & "]<br />"
						'Response.Write "[" & sendSMS_cafe24("[OOO] 신규 온라인상담이 등록되었습니다.(신청자:"&NM_FIELD_1&")", aryRecvSMSList(5, sms_i), "") & "]<br />"
						'Response.End

						Call sendSMS_cafe24("["&m_client_counsel_regist_sms_title&"] 신규 온라인상담이 등록되었습니다("&NM_FIELD_1&"/"&Replace(NM_FIELD_2, "-" ,"")&")", aryRecvSMSList(5, sms_i), "")
					Next
				End If
			End If

			Response.Write "<script type='text/javascript'>alert('온라인상담 등록이 완료되었습니다.');document.location.href='counsel.html';</script>"
		ElseIf MODE = "U" Then
			Response.Write "<script type='text/javascript'>alert('수정되었습니다.');document.location.href='counsel.html';</script>"
		ElseIf MODE = "D" Then
			Response.Write "<script type='text/javascript'>alert('삭제되었습니다.');document.location.href='counsel.html';</script>"
		End If

	Else

		If MODE = "I" Then
			Response.Write "<script type='text/javascript'>alert('온라인상담 등록에 실패하였습니다.\n잠시 후 다시 시도해주세요.');history.back();</script>"
		ElseIf MODE = "U" Then
			Response.Write "<script type='text/javascript'>alert('수정에 실패하였습니다.');document.location.href='counsel.html';</script>"
		ElseIf MODE = "D" Then
			Response.Write "<script type='text/javascript'>alert('삭제에 실패하였습니다.');document.location.href='counsel.html';</script>"
		End If
	End If

	Response.End
%>
