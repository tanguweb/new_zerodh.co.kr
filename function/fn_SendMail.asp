<%
' fn_SendMail(보내는이름, 보내는주소, 받는주소, BodyFormat, MailFormat, 제목, 내용)
Function fn_SendMail(fromName, fromAddr, email, BodyFormat, MailFormat, title, content)

	Dim objConfig ' As CDO.Configuration 
	Dim objMessage ' As CDO.Message 
	Dim Fields ' As ADODB.Fields 

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

	Set objMessage = Server.CreateObject("CDO.Message") 

	Set objMessage.Configuration = objConfig 

		With objMessage 
			.To = email				'받는 사람 주소
			.From = fromName &"<"& fromAddr &">"
			.Subject = title					'메일 제목
			IF MailFormat = 0 Then ' Html형식일때
				.HTMLBody = content
			Else			'Text 형식일때
				.TextBody = content
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

