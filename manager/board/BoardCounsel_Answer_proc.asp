<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_FileSave.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<!-- #include virtual = "/function/fn_Board.asp" -->
<!-- #include virtual = "/function/cafe24/sms/send.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
	'########################################################################################
	'#	File		: /manager/board/BoardDetail_Proc.asp
	'#  Create		: / 2010.09.09
	'#	Info		: 게시판등록
	'#	Update		:
	'#	Update Memo	:
	'########################################################################################


	Dim adoCmd, Result1, Result1_ID, Result1_DEPTH, Result1_SEQ, Result2, Rtn_string
	Dim MODE, CD_BOARDCD, CD_BOARDID, CD_BOARDKEY, NO_BOARD_DEPTH, NO_BOARD_SEQ, YN_USE, NM_TITLE, NM_CONTENTS, NO_CNTVIEW, YN_DISPLAY, NO_VIEWORDER, NO_DOWNLOAD, YN_XML, NM_XMLPATH
	Dim ary_FIELD(16), ary_FIELDTYPE(16)
	Dim UploadForm, arrFileNames(16), arrFileSize(16), arrFileType(16), arrFileTypechk, i, j, iCount, FilePath, FileNameOnly, FileExt, FileName
	Dim HTTP_REFERER, HTTP_USER_AGENT, HTTP_DEVICE_TYPE, HTTP_MEDIA_TYPE

	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
	UploadForm.DefaultPath = FileSaveRoot_Board

	MODE			= UploadForm("MODE")
	'CD_BOARDCD		= UploadForm("CD_BOARDCD")
	CD_BOARDCD		= 1003
	'CD_BOARDID		= UploadForm("CD_BOARDID")

	If MODE = "I" Then
		CD_BOARDID		= 1
	ElseIf MODE = "U" Then
		CD_BOARDID		= UploadForm("answer_CD_BOARDID")
	End If

	'CD_BOARDKEY	= UploadForm("CD_BOARDKEY")
	CD_BOARDKEY		= UploadForm("question_CD_BOARDID")

	'NO_BOARD_DEPTH	= UploadForm("NO_BOARD_DEPTH")
	NO_BOARD_DEPTH	= 2
	'NO_BOARD_SEQ	= UploadForm("NO_BOARD_SEQ")
	NO_BOARD_SEQ	= 2
	'YN_USE			= UploadForm("YN_USE")
	YN_USE			= "Y"
	'NM_TITLE		= UploadForm("NM_TITLE")
	NM_TITLE		= "답변입니다."
	'NM_CONTENTS		= UploadForm("NM_CONTENTS")
	NM_CONTENTS		= Replace(getText(UploadForm("answer_contents")) ,Chr(10),"<br />")
	'NO_CNTVIEW		= UploadForm("NO_CNTVIEW")
	NO_CNTVIEW		= 0
	'YN_DISPLAY		= UploadForm("YN_DISPLAY")
	YN_DISPLAY		= "Y"
	'NO_VIEWORDER	= UploadForm("NO_VIEWORDER")
	NO_VIEWORDER	= 100
	'NO_DOWNLOAD		= UploadForm("NO_DOWNLOAD")
	NO_DOWNLOAD		= 0
	'YN_XML			= UploadForm("YN_XML")
	YN_XML			= ""
	'NM_XMLPATH		= UploadForm("NM_XMLPATH")
	NM_XMLPATH		= ""

for i=1 to 15

	ary_FIELDTYPE(i)	= UploadForm("CD_FIELDTYPE_" & i)
	ary_FIELD(i)		= UploadForm("NM_FIELD_" & i)

next

'	If Mode <> "D" And NM_CONTENTS = "" Then
'		response.write "<script>alert('내용을 입력하세요.');history.back();</script>"
'		response.End
'	End If

	If Mode <> "D" And NM_TITLE = "" Then
		response.write "<script>alert('제목을 입력하세요.');history.back();</script>"
		response.End
	End If



	Dim questionBoard : questionBoard = fn_BoardView(CD_BOARDKEY)
	Dim questionTitle : questionTitle = questionBoard(8, 0)
	Dim questionContents : questionContents = questionBoard(9, 0)
	Dim toSMSphone : toSMSphone = questionBoard(11, 0)			' 상담자의 휴대폰번호
	Dim toEmail : toEmail = questionBoard(12, 0)			' 상담자의 이메일 주소
	Dim isReceiveAnswerEmail : isReceiveAnswerEmail = questionBoard(18, 0)	' 답변 이메일 수신여부



	For j = 1 To 15
		If ary_FIELD(j) <> "" and ary_FIELDTYPE(j) = "FILE" Then
			arrFileNames(j) = fn_FileSave (UploadForm, "NM_FIELD_"&(j), getBoolean(YN_XML="Y" and NM_XMLPATH<>"",FileSaveRoot & NM_XMLPATH,FileSaveRoot_Board))
		End If
	Next

	for i=1 to 15
		if ary_FIELDTYPE(i)="FILE" and UploadForm("NM_FIELD_" & i)="" and UploadForm("hd_NM_FIELD_" & i)<>"" then
			arrFileNames(i)		= UploadForm("hd_NM_FIELD_" & i)
		end if
	next

	'HTTP_REFERER = UploadForm("HTTP_REFERER")
	'HTTP_USER_AGENT = UploadForm("HTTP_USER_AGENT")
	'HTTP_DEVICE_TYPE = UploadForm("HTTP_DEVICE_TYPE")
	'HTTP_MEDIA_TYPE = UploadForm("HTTP_MEDIA_TYPE")

	' 게시글 저장
	SET adoCmd	=	Server.CreateObject("ADODB.Command")
		with adoCmd
		.ActiveConnection	= m_DB
		.CommandType		= adCmdStoredProc
		.CommandText		= "SPA_BOARDDETAIL_EDIT"
		.Parameters.Append	.CreateParameter("@MODE"				, adChar,		adParamInput,	1, Mode)
		.Parameters.Append	.CreateParameter("@CD_BOARDID"			, adInteger,	adParamInput,	 , CD_BOARDID)
		.Parameters.Append	.CreateParameter("@CD_BOARDKEY"			, adInteger,	adParamInput,	 , CD_BOARDKEY)
		.Parameters.Append	.CreateParameter("@NO_BOARD_DEPTH"		, adInteger,	adParamInput,    , NO_BOARD_DEPTH)
		.Parameters.Append	.CreateParameter("@NO_BOARD_SEQ"		, adInteger,	adParamInput,    , NO_BOARD_SEQ)
		.Parameters.Append	.CreateParameter("@CD_BOARDCD"			, adInteger,	adParamInput,    , CD_BOARDCD)
		.Parameters.Append	.CreateParameter("@YN_DISPLAY"			, adchar,		adParamInput,	1, YN_DISPLAY)
		.Parameters.Append	.CreateParameter("@NO_VIEWCNT"			, adInteger,	adParamInput,    , NO_CNTVIEW)
		.Parameters.Append	.CreateParameter("@NO_VIEWORDER"		, adInteger,	adParamInput,	 , NO_VIEWORDER)
		.Parameters.Append	.CreateParameter("@NO_DOWNLOAD"			, adInteger,	adParamInput,	 , NO_DOWNLOAD)
		.Parameters.Append	.CreateParameter("@YN_USE"				, adchar,		adParamInput,	1, YN_USE)
		.Parameters.Append	.CreateParameter("@NM_TITLE"			, adVarchar,	adParamInput,	500, NM_TITLE)
		.Parameters.Append	.CreateParameter("@NM_CONTENTS"			, adVarchar,	adParamInput,	100000, NM_CONTENTS)

If Mode = "D" Then
	for i=1 to 15
		.Parameters.Append	.CreateParameter("@NM_FIELD_"&i			, adVarchar,	adParamInput,1000, "")
	next
Else
	for i=1 to 15
		if  ary_FIELDTYPE(i) = "FILE" Then
			.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, arrFileNames(i))
		else
			.Parameters.Append	.CreateParameter("@NM_FIELD_"&i		, adVarchar,	adParamInput,1000, ary_FIELD(i))
		end if
	next
End If

		.Parameters.Append	.CreateParameter("@CD_USER"				, adVarchar,	adParamInput,	15, m_UserID)
		.Parameters.Append	.CreateParameter("@NO_IPADDR"			, adVarchar,	adParamInput,	20, m_IPAddr)
		.Parameters.Append	.CreateParameter("@HTTP_REFERER"		, adVarchar,	adParamInput,	5000, HTTP_REFERER)
		.Parameters.Append	.CreateParameter("@HTTP_USER_AGENT"		, adVarchar,	adParamInput,	5000, HTTP_USER_AGENT)
		.Parameters.Append	.CreateParameter("@HTTP_DEVICE_TYPE"	, adVarchar,	adParamInput,	200, HTTP_DEVICE_TYPE)
		.Parameters.Append	.CreateParameter("@HTTP_MEDIA_TYPE"		, adVarchar,	adParamInput,	200, HTTP_MEDIA_TYPE)
		.Parameters.Append  .CreateParameter("@Result"				, adVarChar,	adParamoutput,  20	)
		.Parameters.Append  .CreateParameter("@Result_ID"			, adInteger,	adParamoutput  	)
		.Execute
		Result1			= .Parameters("@Result")
		Result1_ID		= .Parameters("@Result_ID")
		end with
	SET adoCmd = Nothing

	if Result1 = "SUCCESS" then

		If MODE = "I" Or MODE = "U" Then
			'----------------------------- start : 이메일 발송 -------------------------------
			Dim isSendMail : isSendMail = m_client_counsel_online_answer_email	'메일 발송할 것인지 여부

			If isSendMail = True Then

				' 상담자가 답변을 이메일로 받는것으로 상담글을 등록한 경우
				If isReceiveAnswerEmail = "Y" Then
					Dim SendMail : SendMail = False
					Dim sFrom, sTo, objMail, MailConfig
					'sFrom = ""&NM_NAME&"" & "<webmaster@domain>" ' 보내는사람
					sFrom = "<"&CONFIG_SMTP_SENDUSERNAME&">" ' 보내는사람
					sTo = toEmail
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
					objMail.Subject = "["&m_client_counsel_regist_email_title&"] 온라인상담 답변 입니다."

					Dim mailBody

					'mailBody = mailBody & "<strong>* 이름 : </strong>" & NM_FIELD_1 & "<br /><br />"
					'mailBody = mailBody & "<strong>* 전화번호 : </strong>" & NM_FIELD_2 & "<br /><br />"
					'mailBody = mailBody & "<strong>* 휴대폰번호 : </strong>" & NM_FIELD_3 & "<br /><br />"
					'mailBody = mailBody & "<strong>* 이메일 : </strong>" & NM_FIELD_4 & "<br /><br />"
					''mailBody = mailBody & "<strong>* 행사예정일 : </strong>" & NM_FIELD_5 & "<br /><br />"
					''mailBody = mailBody & "<strong>* 행사예정시간 : </strong>" & NM_FIELD_6 & "<br /><br />"
					''mailBody = mailBody & "<strong>* 예상인원수 : </strong>" & NM_FIELD_7 & "<br /><br />"
					''mailBody = mailBody & "<strong>* 식사가격대 : </strong>" & NM_FIELD_8 & "<br /><br />"
					'mailBody = mailBody & "<strong>* 내용 : </strong>" & NM_CONTENTS & "<br /><br />"


mailBody = mailBody & "<html lang='ko'><head>"
mailBody = mailBody & "    <title>천안 "&m_client_name&"</title>"
mailBody = mailBody & "    <meta charset='UTF-8'>"
mailBody = mailBody & "    <meta http-equiv='X-UA-Compatible' content='IE=edge'>"
mailBody = mailBody & "    <meta name='viewport' content=''>"
mailBody = mailBody & "    <meta name='author' content=''>"
mailBody = mailBody & "    <meta name='keywords' content=''>"
mailBody = mailBody & "    <meta name='description' content=''>"
mailBody = mailBody & "    <meta name='publisher' content=''>"
mailBody = mailBody & "    <link rel='stylesheet' type='text/css' href='https://fonts.googleapis.com/css?family=Montserrat:100,200,300,400,500,600,700,800,900|Nanum+Gothic:400,700,800|Noto+Sans+KR:100,300,400,500,700,900|Play:400,700&amp;subset=korean'>"
mailBody = mailBody & "  </head>"
mailBody = mailBody & "  <body naver_screen_capture_injected='true'>"
mailBody = mailBody & "  <div style='"
mailBody = mailBody & "        width: 630px;"
mailBody = mailBody & "        margin: 0 auto;"
mailBody = mailBody & "        border: 1px solid #ccc;"
mailBody = mailBody & "        color: #666;"
mailBody = mailBody & "        z-index: -10;"
mailBody = mailBody & "        '>"
mailBody = mailBody & "    <h1 style='text-align: center;font-size: 24px;padding: 20px 0 20px;margin:0;font-family: Noto Sans KR;>"
'mailBody = mailBody & "    <img src='http://cheonwgj.co.kr/html/images/common/logo.png' alt='' style='margin: 0 0 20px 0; width:200px;'><br>"
'mailBody = mailBody & "    <img src='http://cheonwgj.co.kr/html/introduce/images/photo_01.jpg' alt='' style='width:630px;'>"
mailBody = mailBody & "    <p style='font-weight: 500;'>“ 안녕하세요 "&m_client_name&"입니다”</p></h1>"
mailBody = mailBody & "      <div style='border: 1px solid  #396ebb;border-radius: 4px;margin: 00px 20px 20px 20px;padding: 0 0 15px;position: relative;overflow: hidden;min-height: 150px;background: #fff;'>"
mailBody = mailBody & "              <h5 style='font-size: 16px;margin-bottom: 10px;background: #396ebb;padding: 10px 15px;margin: 0;color: #fff;'>"&questionTitle&"</h5>"
mailBody = mailBody & "              <div style='padding: 15px;z-index: 1;position: relative;color: #777;line-height: 24px;'>"
mailBody = mailBody & questionContents	' 질문내용
mailBody = mailBody & "        </div>"
mailBody = mailBody & "              <span style='position: absolute;font-size: 100px;bottom: -10px;right: 0;color: rgba(57, 110, 187, .2);font-family: tahoma;font-weight: 900;z-index: 0;font-family: Play, sans-serif;'>Q</span>"
mailBody = mailBody & "            </div>"
mailBody = mailBody & "            <div style='border: 1px solid #8b8b8b;border-radius: 4px;margin:10px 20px;padding: 15px;position: relative;overflow: hidden;margin-bottom: 20px;min-height: 130px;background: #fff;color: #777;'>"
mailBody = mailBody & "              <div style='position: relative;z-index: 1;line-height: 24px;'>"
mailBody = mailBody & NM_CONTENTS	' 답변내용
mailBody = mailBody & "              </div>"
mailBody = mailBody & "              <span style='position: absolute;font-size: 100px;bottom: -20px;color:rgba(156, 156, 156, 0.2);right: 0;font-family: tahoma;font-weight: 900;z-index: 0;font-family: Play, sans-serif;'>A</span>"
mailBody = mailBody & "            </div>"
mailBody = mailBody & "            <div style='text-align: center;margin: 0 20px 0;padding-top: 10px;'>"
mailBody = mailBody & "              <a href='"&m_client_domain&"/html/board/sub_01.html' target='_blank'><span style='background: #8e8e8e;padding: 10px 20px;display: inline-block;margin-bottom: 30px;border-radius: 4px;color: #fff;letter-spacing: -1px;width: 120px;'>온라인 상담</span></a>"
mailBody = mailBody & "              <a href='"&m_client_domain&"' target='_blank'><span style='background: #396ebb;padding: 10px 20px;display: inline-block;margin-bottom: 30px;border-radius: 4px;color: #fff;width: 120px;margin-left: 10px;'>홈페이지</span></a>"
mailBody = mailBody & "            </div>"
mailBody = mailBody & "            <p style='background: #282828;color: #a6a6a6; font-size: 13px;margin:0;padding: 10px  20px; text-align: center;'>* 본 메일은 발신전용으로 회신되지 않습니다. 문의 사항은 고객센터로 연락 주시기 바랍니다.</p>"
mailBody = mailBody & "    <div style='background: #505050;color: #ccc;font-style: normal;padding: 15px  20px;font-size: 12px; line-height: 16px; text-align: center;'>"
mailBody = mailBody & "          <span>"&m_client_name&"의원 </span>"
mailBody = mailBody & "          <span style='margin-left:15px'>"&m_client_address1&" </span><span style='margin-left:15px'>Tel.   "&m_client_tel&"   </span>  <br>"
mailBody = mailBody & "          <span>대표자명 : "&m_client_owner&"</span><span style='margin-left:15px'>사업자등록번호 :  "&m_client_business_no&"</span>  <br>"
mailBody = mailBody & "          <span>COPYRIGHT © 2020 "&m_client_name&"의원  ALL RIGHT RESERVED.</span>"
mailBody = mailBody & "    </div>"
mailBody = mailBody & "  </div>"
mailBody = mailBody & "</body></html>"


					'response.write "mailBody : " & mailBody

					objMail.HTMLBody = mailBody

					objMail.BodyPart.Charset="utf-8"
					objMail.HTMLBodyPart.Charset="utf-8"

					objMail.send

					Set objMail = Nothing
					If Not Err Then
						SendMail = True
					Else
						SendMail = False
					End If
				End If

			End If
			'----------------------------- end : 이메일 발송 -------------------------------



			Dim isSendSMS : isSendSMS = m_client_counsel_online_answer_sms	'SMS 발송할 것인지 여부

			If isSendSMS = True Then
				If toSMSphone <> "" Then
					Call sendSMS_cafe24("["&m_client_counsel_regist_sms_title&"] 문의해주신 온라인 상담글에 답변이 등록되었습니다.^^", addHyphen(Replace(toSMSphone,"-","")), "")
				End If
			End If



		End If



		If MODE = "I" Then	' 등록
			Rtn_string = "<script>alert('답변이 저장되었습니다.');parent.document.location.reload();</script>"
		ElseIf MODE = "U" Then	' 수정
			Rtn_string = "<script>alert('답변이 수정되었습니다.');parent.document.location.reload();</script>"
		ElseIf MODE = "D" Then	' 삭제
			Rtn_string = "<script>alert('답변이 삭제되었습니다.');parent.document.location.reload();</script>"
		End If
		response.write Rtn_string





		response.end
	else
		response.write "<script>alert('답변이 정상적으로 등록되지 못하였습니다.');</script>"
		response.end
	end if

%>
