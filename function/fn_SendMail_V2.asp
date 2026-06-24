<%
' #### 이메일 발송 함수 ######################
	Public Function SendMail(strSenderAdd , strReceiverAdd, strSubject , strMailBody)
		SendMail = False
		Dim sFrom, sTo, objMail, MailConfig
		sFrom = strSenderAdd ' 보내는사람
		sTo = strReceiverAdd ' 받는사람
		Set objMail = Server.CreateObject("CDO.Message") ' CDO 2.0(메일 보내기 컴포넌트 개체 생성)

		Set MailConfig = objMail.Configuration
		With MailConfig.Fields
		
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
		Set MailConfig=Nothing

		objMail.From = sFrom ' 메일을 보내는 사람의 이메일 주소
		objMail.To = sTo ' 메일을 받는 사람의 이메일주소(여러사람일 경우는 ; 표시로 구분)
		objMail.Subject = strSubject ' 메일 제목
		objMail.HTMLBody = strMailBody
		objMail.BodyPart.Charset = "ks_c_5601-1987"     
		objMail.HTMLBodyPart.Charset = "ks_c_5601-1987"

		objMail.send
		Set objMail = Nothing
		If Not Err Then
			SendMail = True
		Else
			SendMail = False
		End If
	End Function
%>
