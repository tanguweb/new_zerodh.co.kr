<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Type Library" -->

<%
Class clsCDO
	private objCDO
	private objConfig


	private m_SMTP
	private m_From
	private m_To
	private m_Subject
	private m_Body
	private m_Type
	private m_UserName
	private m_Password


	Private Sub Class_Initialize
		Set objCDO = Server.CreateObject("CDO.Message")
		Set objConfig = Server.CreateObject("CDO.Configuration")

		setSMTPServer("192.168.1.2")
		m_Type = "HTML"
	End Sub

	Private Sub Class_Terminate
		Set objCDO = Nothing
		Set objConfig = nothing
	End Sub

	Private Sub setSMTPServer(strSMTP)
		Set Flds = objConfig.Fields
		'		1 : cdoSendUsingPickup
		'		2 : cdoSendUsingPort

		' ## Schema 사용시 (typelib 로드 필요움따)
		'Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		'Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
		'Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

		' ## 상수 사용시 (typelib 로드가 필요하다)
		Flds.Item(cdoSendUsingMethod) = cdoSendUsingPort
		Flds.Item(cdoSMTPServer) = strSMTP
		Flds.Item(cdoSMTPServerPort) = 25
		If m_UserName <> "" And m_Password <> "" Then
		Flds.Item(cdoSMTPAuthenticate) = cdoBasic
		Flds.Item(cdoSendUserName) = m_UserName
		Flds.Item(cdoSendPassword) = m_Password
		End If


		Flds.update
		Set objCDO.Configuration = objConfig
		Set Flds = Nothing
	End Sub




	Public Property let SMTP(p_SMTP)
		setSMTPServer(p_SMTP)
		m_SMTP = p_SMTP
	End Property
	Public Property get SMTP()
		SMTP = m_SMTP
	End Property


	Public Property let MailFrom(p_From)
		m_From = p_From
	End Property
	Public Property get MailFrom()
		From = m_From
	End Property


	Public Property let MailTo(p_To)
		m_To = p_To
	End Property
	Public Property get MailTo()
		MailTo = m_To
	End Property


	Public Property let Subject(p_Subject)
		m_Subject = p_Subject
	End Property
	Public Property get Subject()
		Subject = m_Subject
	End Property


	Public Property let Body(p_Body)
		m_Body = p_Body
	End Property
	Public Property get Body()
		Body = m_Body
	End Property


	Public Property let MailType(p_Type)
		m_Type = p_Type
	End Property
	Public Property get MailType()
		MailType = m_Type
	End Property


	Public Property let UserName(p_UserName)
		m_UserName = p_UserName
	End Property
	Public Property get UserName()
		UserName = m_UserName
	End Property


	Public Property let Password(p_Password)
		m_Password = p_Password
	End Property
	Public Property get Password()
		Password = m_Password
	End Property


	Public Sub Send()
		objCDO.MimeFormatted = false

		objCDO.From = m_From
		objCDO.To = m_To
		objCDO.Subject = m_Subject

		If Ucase(m_Type) = "TEXT" Then
		objCDO.TextBody = m_Body
		objCDO.BodyPart.Charset = "gb2312"
		Else
		objCDO.HTMLBody = m_Body
		objCDO.HTMLBodyPart.Charset = "gb2312"
		End If

		objCDO.Send
	End Sub

End Class


'set obj = new clsCDO
'	obj.MailFrom = "goonis2@nate.com"
'	obj.MailTo = "goonis2@nate.com"
'	obj.Subject = "test"
'	obj.Body = "test cdo"

'	obj.Send()
'set obj = nothing


%>