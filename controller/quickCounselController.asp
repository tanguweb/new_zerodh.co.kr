<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/JSON_2.0.4.asp" -->
<!-- #include virtual = "/function/JSON_UTIL_0.1.1.asp" -->
<!-- #include virtual = "/function/cafe24/sms/send.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp"-->
<%
	'########################################################################################
	'#	File		: /controller/mainCounselController.asp
	'#  Create		: 2015.11
	'#	Info		: 
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	
	Dim CD_BOARDCD : CD_BOARDCD = 1003
	Dim objADO, SQL, aryLoginInfo
	Dim adoCmd, Result1, Result1_ID
	Dim ary_FIELD(16), i
	Dim MODE : MODE = "I"
	Dim CD_BOARDID : CD_BOARDID = 0

	Dim NM_FIELD_1  : NM_FIELD_1 = getText(ReqF("quickCounselUserName"))
	Dim NM_FIELD_2  : NM_FIELD_2 = getText(ReqF("quickCounselPhone"))
	Dim NM_FIELD_4  : NM_FIELD_4 = "9999"	'"빠른 상담"	'getText(UploadForm("counselType"))
	Dim NM_FIELD_7  : NM_FIELD_7 = m_client_counsel_pw	'getText(UploadForm("password"))
	Dim NM_FIELD_9  : NM_FIELD_9 = "N" 'getText(UploadForm("mail_ok"))
	Dim NM_FIELD_11 : NM_FIELD_11 = "READY"
	'Dim NM_TITLE	: NM_TITLE = "빠른 상담 신청입니다. 신청자: " & NM_FIELD_1 & " 님."
	Dim NM_TITLE	: NM_TITLE = (Left(Trim(NM_FIELD_1), Len(Trim(NM_FIELD_1)) - 1) & "*") & "님이 빠른 상담을 남기셨습니다."
	'Dim NM_TITLE	: NM_TITLE = maskingString(Trim(NM_FIELD_1)) & "님이 빠른 상담을 남기셨습니다."

	Dim NM_CONTENTS : NM_CONTENTS = Replace(getText(ReqF("quickCounselContent")) ,Chr(13)&Chr(10),"<br />")
	Dim HTTP_REFERER : HTTP_REFERER = ReqF("HTTP_REFERER")
	Dim HTTP_USER_AGENT : HTTP_USER_AGENT = ReqF("HTTP_USER_AGENT")
	Dim HTTP_DEVICE_TYPE : HTTP_DEVICE_TYPE = ReqF("HTTP_DEVICE_TYPE")
	Dim HTTP_MEDIA_TYPE : HTTP_MEDIA_TYPE = ReqF("HTTP_MEDIA_TYPE")

	'json 객체 선언
	Dim jsonResult 
	
	Set jsonResult = jsObject()
	jsonResult("result") = "ERROR"

	If NM_FIELD_1 = "" Or NM_FIELD_2 = "" Or NM_CONTENTS = "" Then 
		Response.Write toJSON(jsonResult)
		Response.End
	End If 

	For i = 1 To 15
		ary_FIELD(i)		= ReqF("NM_FIELD_" & i)
	Next

	ary_FIELD(1) = NM_FIELD_1
	ary_FIELD(2) = NM_FIELD_2
	ary_FIELD(4) = NM_FIELD_4
	ary_FIELD(7) = NM_FIELD_7
	ary_FIELD(9) = NM_FIELD_9
	ary_FIELD(11) = NM_FIELD_11
	
	' 게시글 저장
	Set adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDDETAIL_EDIT"
		.Parameters.Append	.CreateParameter("@MODE"			, adChar,		adParamInput,	1, MODE)
		.Parameters.Append	.CreateParameter("@CD_BOARDID"		, adInteger,	adParamInput,	 , CD_BOARDID)
		.Parameters.Append	.CreateParameter("@CD_BOARDKEY"		, adInteger,	adParamInput,	 , 0)
		.Parameters.Append	.CreateParameter("@NO_BOARD_DEPTH"	, adInteger,	adParamInput,    , 1)
		.Parameters.Append	.CreateParameter("@NO_BOARD_SEQ"	, adInteger,	adParamInput,    , 1)
		.Parameters.Append	.CreateParameter("@CD_BOARDCD"		, adInteger,	adParamInput,    , CD_BOARDCD)
		.Parameters.Append	.CreateParameter("@YN_DISPLAY"		, adchar,		adParamInput,	1, "Y")
		.Parameters.Append	.CreateParameter("@NO_VIEWCNT"		, adInteger,	adParamInput,    , 0)
		.Parameters.Append	.CreateParameter("@NO_VIEWORDER"	, adInteger,	adParamInput,	 , 100)
		.Parameters.Append	.CreateParameter("@NO_DOWNLOAD"		, adInteger,	adParamInput,	 , 0)
		.Parameters.Append	.CreateParameter("@YN_USE"			, adchar,		adParamInput,	1, "Y")
		.Parameters.Append	.CreateParameter("@NM_TITLE"		, adVarchar,	adParamInput,	500, NM_TITLE)
		.Parameters.Append	.CreateParameter("@NM_CONTENTS"		, adVarchar,	adParamInput,	100000, NM_CONTENTS)
	For i = 1 To 15
		.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, ary_FIELD(i))
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
		
		Dim isSendMail : isSendMail = m_client_counsel_footer_send_email 	'메일 발송할 것인지 여부
		
		If isSendMail = True Then 
			Dim SendMail : SendMail = False
			Dim sFrom, sTo, objMail, MailConfig
			'sFrom = ""&NM_NAME&"" & "<webmaster@domain>" ' 보내는사람
			sFrom = "<"&CONFIG_SMTP_SENDUSERNAME&">" ' 보내는사람
			'sTo = "<master@peters.co.kr>"	' 받는사람
			sTo = ""
			
			Dim aryRecvEMailList, email_i
			aryRecvEMailList = fn_CodeInfoWhere("COUNSEL_MAIL_RECV", " AND NM_CODE1 = '"&CD_BOARDCD&"'")
			
			'Response.write "IsArray(aryRecvEMailList) : " & IsArray(aryRecvEMailList) & "<br />"

			If IsArray(aryRecvEMailList) Then 
				For email_i = 0 To UBound(aryRecvEMailList, 2)
					
					'Response.write "aryRecvEMailList(5, "&email_i&") : " & aryRecvEMailList(5, email_i) & "<br />"

					If sTo = "" Then
						sTo = sTo & "<" & aryRecvEMailList(5, email_i) & ">"
						'sTo = sTo & "" & aryRecvEMailList(5, email_i) & ""
						'Response.write ""&email_i&" sTo assign : " & "" & aryRecvEMailList(5, email_i) & "" & "<br />"
						'Response.write ""&email_i&" sTo : " & sTo & "<br />"
					Else 
						sTo = sTo & ",<" & aryRecvEMailList(5, email_i) & ">"
						'sTo = sTo & "," & aryRecvEMailList(5, email_i) & ""
						'Response.write ""&email_i&" sTo : " & sTo & "<br />"
					End If 
				Next 
			End If
			
			'Response.write "sTo : " & sTo & "<br />"
			
			If sTo <> "" Then 
				'Response.write "sTo 123 : " & sTo & "<br />"
				
				'jsonResult("sTo") = sTo

				Set objMail = Server.CreateObject("CDO.Message") ' CDO 2.0(메일 보내기 컴포넌트 개체 생성)
				

				Set MailConfig = objMail.Configuration
				With MailConfig.Fields
				.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2	' 1 : 로컬 SMTP) / 2 : 외부 SMTP)
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") =  "C:\Inetpub\mailroot\Pickup"	' Pickup 디렉토리 설정
				'.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"  ' 호스트설정
				.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mw-002.cafe24.com"
				.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25	' SMTP Port
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" ) = 1 
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusername" ) = CONFIG_SMTP_SENDUSERNAME
				.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = CONFIG_SMTP_SENDPASSWORD	'(리얼)
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30	' 연결시간  
				.Update
				End With
				Set MailConfig = Nothing

				objMail.From = sFrom ' 메일을 보내는 사람의 이메일 주소
				objMail.To = sTo ' 메일을 받는 사람의 이메일주소(여러사람일 경우는 ; 표시로 구분)
				'objMail.Subject = NM_FIELD_1 & "상담" ' 메일 제목
				objMail.Subject = "["&m_client_counsel_regist_email_title&"] 빠른 상담 신청이 접수되었습니다." & " (신청자: " & NM_FIELD_1 & ")"
				
				Dim mailBody

				mailBody = mailBody & "<strong>* 이름 : </strong>" & NM_FIELD_1 & "<br /><br />"
				mailBody = mailBody & "<strong>* 연락처 : </strong>" & NM_FIELD_2 & "<br /><br />"
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

				jsonResult("email_result") = SendMail
			End If 

		End If
		
		Dim isSendSMS : isSendSMS = m_client_counsel_footer_send_sms 	'SMS 발송할 것인지 여부
		
		If isSendSMS = True Then

			Dim aryRecvSMSList, sms_i
			aryRecvSMSList = fn_CodeInfoWhere("COUNSEL_SMS_RECEIVER", " AND NM_CODE1 = '"&CD_BOARDCD&"'")
			
			If IsArray(aryRecvSMSList) Then 
				For sms_i = 0 To UBound(aryRecvSMSList, 2)
					jsonResult("sms_result") = sendSMS_cafe24("["&m_client_counsel_regist_sms_title&"] 신규 간편상담이 등록되었습니다("&NM_FIELD_1&"/"&Replace(NM_FIELD_2, "-" ,"")&")", aryRecvSMSList(5, sms_i), "")
				Next 
			End If
		End If

		'Response.Write "<script type='text/javascript'>alert('업무제휴 문의가 완료되었습니다.');document.location.href='business.asp';</script>"
		jsonResult("result") = Result1
	Else
		'Response.Write "<script type='text/javascript'>alert('업무제휴 문의에 실패하였습니다.\n잠시 후 다시 시도해주세요.');history.back();</script>"
		jsonResult("result") = Result1
	End If 

	'Response.End 
	
	'jsonResult.Flush
	Response.Write toJSON(jsonResult)
	Response.End 
%>