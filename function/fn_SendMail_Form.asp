<%
' fn_SendMail(보내는이름, 보내는주소, 받는주소, BodyFormat, MailFormat, 제목, 내용)
Function fn_SendMail_Form(fromName, fromAddr, email, BodyFormat, MailFormat, title, content)

	Dim objConfig ' As CDO.Configuration 
	Dim objMessage ' As CDO.Message 
	Dim Fields ' As ADODB.Fields 
	Dim strMailBody

	' Get a handle on the config object and it's fields 
	Set objConfig = Server.CreateObject("CDO.Configuration") 
	Set Fields = objConfig.Fields 
	
	' Set config fields we care about 
	With Fields 
		If CONFIG_SMTP_SENDUSING = 1 Then 
			.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = CONFIG_SMTP_SENDUSING	' 1 : 로컬 SMTP) / 2 : 외부 SMTP)
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") =  CONFIG_SMTP_SMTPSERVERPICKUPDIRECTORY	' Pickup 디렉토리 설정
			.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CONFIG_SMTP_SMTPSERVER		' 호스트설정
			.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CONFIG_SMTP_SMTPSERVERPORT  
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = CONFIG_SMTP_SMTPCONNECTIONTIMEOUT   
		ElseIf CONFIG_SMTP_SENDUSING = 2 Then 
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = CONFIG_SMTP_SENDUSING
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CONFIG_SMTP_SMTPSERVER
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = CONFIG_SMTP_SMTPSERVERPORT 
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" ) = CONFIG_SMTP_SMTPAUTHENTICATE
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusername" ) = CONFIG_SMTP_SENDUSERNAME
			.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = CONFIG_SMTP_SENDPASSWORD
		End If

		.Update 
	End With 
	
	strMailBody = ""
	strMailBody = strMailBody & "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>" & vbcrlf
	strMailBody = strMailBody & "<html xmlns='http://www.w3.org/1999/xhtml'>" & vbcrlf
	strMailBody = strMailBody & "<head>" & vbcrlf
	strMailBody = strMailBody & "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />" & vbcrlf
	strMailBody = strMailBody & "<title>무제 문서</title>" & vbcrlf
	strMailBody = strMailBody & "<style type='text/css'>" & vbcrlf
	strMailBody = strMailBody & ".top {width:696px; height:auto; " & vbcrlf
	strMailBody = strMailBody & "	border-top: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "   border-left: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "   border-right: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "	position:relative; left:50%; margin-left:-348px; " & vbcrlf
	strMailBody = strMailBody & "}" & vbcrlf
	strMailBody = strMailBody & ".content {width:696px; height:auto; " & vbcrlf
	strMailBody = strMailBody & "   border-left: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "   border-right: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "	position:relative; left:50%; " & vbcrlf
	strMailBody = strMailBody & "	margin:0 0 0 -348px;" & vbcrlf
	strMailBody = strMailBody & "}" & vbcrlf
	strMailBody = strMailBody & ".buttom {width:696px; height:auto; " & vbcrlf
	strMailBody = strMailBody & "	border-buttom: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "   border-left: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "   border-right: #e5e5e5 1px solid ;" & vbcrlf
	strMailBody = strMailBody & "	position:relative; left:50%; margin-left:-348px; " & vbcrlf
	strMailBody = strMailBody & "}" & vbcrlf

	strMailBody = strMailBody & "body,td,th {" & vbcrlf
	strMailBody = strMailBody & "	font-family: Verdana, Geneva, sans-serif;" & vbcrlf
	strMailBody = strMailBody & "	font-size: 12px;" & vbcrlf
	strMailBody = strMailBody & "	color: #999;" & vbcrlf
	strMailBody = strMailBody & "}" & vbcrlf
	strMailBody = strMailBody & "</style>" & vbcrlf
	strMailBody = strMailBody & "</head>" & vbcrlf

	strMailBody = strMailBody & "<body>" & vbcrlf
	strMailBody = strMailBody & "<div class='top'>" & vbcrlf
	strMailBody = strMailBody & "<img src='http://www.11dent.co.kr/images/mail/mail_form_top.jpg' width='696' height='81' border='0' usemap='#Map' />" & vbcrlf
	strMailBody = strMailBody & "<map name='Map' id='Map'>" & vbcrlf
	strMailBody = strMailBody & "			<area shape='rect' coords='307,45,393,63' href='http://www.11dent.co.kr/_company/company.asp' target='_blank' alt='일레븐치과 소개' />" & vbcrlf
	strMailBody = strMailBody & "			<area shape='rect' coords='406,46,491,63' href='http://www.11dent.co.kr/_special/special.asp' target='_blank' />" & vbcrlf
	strMailBody = strMailBody & "			<area shape='rect' coords='503,45,600,64' href='http://www.11dent.co.kr/_quickort/quickort.asp' target='_blank' />" & vbcrlf
	strMailBody = strMailBody & "			<area shape='rect' coords='614,45,688,63' href='http://www.11dent.co.kr/FriendlyConsultation/TwentyFour_List.asp' target='_blank' />" & vbcrlf
	strMailBody = strMailBody & "			<area shape='rect' coords='14,4,178,64' href='http://www.11dent.co.kr/' target='_blank' alt='일레븐치과' />" & vbcrlf
	strMailBody = strMailBody & "</map>" & vbcrlf
	strMailBody = strMailBody & "</div >" & vbcrlf

	strMailBody = strMailBody & "<div class='content'>" & vbcrlf
	strMailBody = strMailBody & content
	strMailBody = strMailBody & "</div>" & vbcrlf
	strMailBody = strMailBody & "<div class='buttom'>" & vbcrlf
	strMailBody = strMailBody & "<img src='http://www.11dent.co.kr/images/mail/mail_form_buttom.jpg' width='696' height='95' border='0' usemap='#Map2' />" & vbcrlf
	strMailBody = strMailBody & "<map name='Map2' id='Map2'>" & vbcrlf
	strMailBody = strMailBody & "			<area shape='rect' coords='13,15,159,78' href='http://www.11dent.co.kr/' target='_blank' alt='일레븐치과' />" & vbcrlf
	strMailBody = strMailBody & "</map>" & vbcrlf
	strMailBody = strMailBody & "</div>" & vbcrlf
	strMailBody = strMailBody & "</body>" & vbcrlf
	strMailBody = strMailBody & "</html>"

	Set objMessage = Server.CreateObject("CDO.Message") 

	Set objMessage.Configuration = objConfig 

		With objMessage 
			.To = email				'받는 사람 주소
			'.From = fromName &"<"& fromAddr &">"
			.From = "<"& fromAddr &">"
			.Subject = title					'메일 제목
			IF MailFormat = 0 Then ' Html형식일때
				.HTMLBody = strMailBody
			Else			'Text 형식일때
				.TextBody = strMailBody
			End If    
			.BodyPart.Charset = "ks_c_5601-1987"     
			.HTMLBodyPart.Charset = "ks_c_5601-1987"
			.Send 
		End With 

'	Response.Write "Success" 

	Set Fields = Nothing 
	Set objMessage = Nothing 
	Set objConfig = Nothing 
	 
End Function
%>

