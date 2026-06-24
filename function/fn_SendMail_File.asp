<%
' fn_SendMail(보내는이름, 보내는주소, 받는주소, BodyFormat, MailFormat, 제목, 내용)
Function fn_SendMail_File_3(fromName, fromAddr, email, BodyFormat, MailFormat, title, content, strFileObj1, strFileObj2, strFileObj3)

	Dim objConfig ' As CDO.Configuration 
	Dim objMessage ' As CDO.Message 
	Dim Fields ' As ADODB.Fields 
	Dim strMailBody

	' Get a handle on the config object and it's fields 
	Set objConfig = Server.CreateObject("CDO.Configuration") 
	Set Fields = objConfig.Fields 
	
	' Set config fields we care about 
	With Fields 
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = CONFIG_SMTP_SENDUSING	'내부
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") =  CONFIG_SMTP_SMTPSERVERPICKUPDIRECTORY
		'.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CONFIG_SMTP_SMTPSERVER  '테스트
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = CONFIG_SMTP_SMTPSERVERPORT 
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" ) = CONFIG_SMTP_SMTPAUTHENTICATE 
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = CONFIG_SMTP_SMTPCONNECTIONTIMEOUT 

		'리얼 발송
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2	'외부
		''.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1	'내부
		''.Item(cdoSMTPServer) = "mw-001.cafe24.com"	' <-- 여기는 발송이 안되는거 같습니다.
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mw-002.cafe24.com" '리얼
		''.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1" '리얼
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = 25 
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" ) = 1 
		''.Item("http://schemas.microsoft.com/cdo/configuration/sendusername" ) = fromAddr '보내는 주소(리얼)
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendusername" ) = "webmaster@thementor7.cafe24.com"
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "dkssud123"	'(리얼)
		
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = CONFIG_SMTP_SENDUSING
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CONFIG_SMTP_SMTPSERVER
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = CONFIG_SMTP_SMTPSERVERPORT 
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" ) = CONFIG_SMTP_SMTPAUTHENTICATE
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusername" ) = CONFIG_SMTP_SENDUSERNAME
		.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = CONFIG_SMTP_SENDPASSWORD

		.Update 
	End With 
	
	strMailBody = ""
	strMailBody = strMailBody & "<!DOCTYPE html>"
	strMailBody = strMailBody & "<html xmlns='http://www.w3.org/1999/xhtml' xml:lang='ko' lang='ko'>"
	strMailBody = strMailBody & "<head>"
	strMailBody = strMailBody & "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />"
	strMailBody = strMailBody & "<title></title>"
	strMailBody = strMailBody & "</head>"
	strMailBody = strMailBody & "<body>"
	strMailBody = strMailBody & "	<div class='mento_mail' style='margin:30px auto 0 auto;padding:0;width:720px;'>"
	strMailBody = strMailBody & "		<div style='margin:0;padding:0;font-size:0'><img src='" & m_SiteDomain & "images/content/mail_top.gif' border='0'/></div>"
	strMailBody = strMailBody & "		<div style='margin:0;padding:50px 39px 80px 39px;border-left:1px solid #d3d3d3;border-right:1px solid #d3d3d3;background:#fff'>"
	'strMailBody = strMailBody & "			<p style='margin:0;padding:0;text-align:center'><img src='" & m_SiteDomain & "images/content/mail_request_txt.gif' alt='더멘토 성형외과의 가족이 되신 것을 진심으로 환영합니다.' border='0'/></p>"
	'strMailBody = strMailBody & "			<p style='margin:0;padding:0 15px;margin-top:30px;font-size:14px;height:39px;line-height:39px;text-align:left;border:1px solid #abb9cc;background:#dde3ec;'>"
	'strMailBody = strMailBody & "			<strong>" & NM_NAME & "</strong>님의 온라인 상담내용"
	'strMailBody = strMailBody & "			</p>"
	strMailBody = strMailBody & "			<div style='margin:0;padding:30px 15px 20px 15px;font-size:12px;text-align:left;color:#888;line-height:18px;'>"
	strMailBody = strMailBody & Replace(content,Chr(13)&Chr(10),"<br />")
	strMailBody = strMailBody & "			</div>"
	'strMailBody = strMailBody & "			<div style='margin:0;padding:20px 25px 20px 25px;font-size:12px;text-align:left;color:#888;line-height:18px;border-bottom:2px solid #a6a6a6;background:#fafafa'>"
	'strMailBody = strMailBody & "				<p style='margin:0;padding:0;font-size:14px;color:#000;margin-bottom:10px'>답변입니다.</p>"
	'strMailBody = strMailBody & NM_ANSWER
	'strMailBody = strMailBody & "			</div>"
	strMailBody = strMailBody & "		</div>"
	strMailBody = strMailBody & "		<div style='margin:0;padding:0;font-size:0'>"
	strMailBody = strMailBody & "			<img src='" & m_SiteDomain & "images/content/mail_customer.gif' border='0' usemap='#Map' />"
	strMailBody = strMailBody & "			<map name='Map' id='Map'>"
	strMailBody = strMailBody & "			  <area shape='rect' coords='408,23,516,160' href='" & m_SiteDomain & "advice/advice_reservation.asp' target='_blank' title='온라인예약'/>"
	strMailBody = strMailBody & "			  <area shape='rect' coords='580,24,686,167' href='" & m_SiteDomain & "advice/advice_list.asp' target='_blank' title='온라인상담'/>"
	strMailBody = strMailBody & "			  <area shape='rect' coords='393,194,545,249' href='" & m_SiteDomain & "' target='_blank' title='홈페이지 바로가기'/>"
	strMailBody = strMailBody & "			</map>"
	strMailBody = strMailBody & "		</div>"
	strMailBody = strMailBody & "		<div style='margin:0;padding:0;font-size:0'><img src='" & m_SiteDomain & "images/content/mail_bottom.gif' border='0'/></div>"
	strMailBody = strMailBody & "	</div>"
	strMailBody = strMailBody & "</body>"
	strMailBody = strMailBody & "</html>"


	'Response.Write getText(strMailBody)
	'Response.End 
	Set objMessage = Server.CreateObject("CDO.Message") 

	Set objMessage.Configuration = objConfig 

		With objMessage 
			.To = email				'받는 사람 주소
			.From = fromName &"<"& fromAddr &">"
			'.From = "<"& fromAddr &">"
			.Subject = title					'메일 제목
			IF MailFormat = 0 Then ' Html형식일때
				.HTMLBody = strMailBody
			Else			'Text 형식일때
				.TextBody = strMailBody
			End If
			.BodyPart.Charset = "ks_c_5601-1987"     
			.HTMLBodyPart.Charset = "ks_c_5601-1987"
			If strFileObj1 <> "" Then
				strFileObj1 = Replace(strFileObj1,"/","\")
				.AddAttachment(strFileObj1)
			End If
			If strFileObj2 <> "" Then
				strFileObj2 = Replace(strFileObj2,"/","\")
				.AddAttachment(strFileObj2)
			End If
			If strFileObj3 <> "" Then
				strFileObj3 = Replace(strFileObj3,"/","\")
				.AddAttachment(strFileObj3)
			End If
			.Send 
		End With 

'	Response.Write "Success" 

	Set Fields = Nothing 
	Set objMessage = Nothing 
	Set objConfig = Nothing 
	 
End Function
%>

