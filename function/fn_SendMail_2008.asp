<%
' ######## 메일 보내기 설정
Public Function SendMail(to_email , from_email, subject, content)

	'response.write to_email & "<br>"
	'response.write from_email & "<br>"
	'response.write subject & "<br>"
	'response.write content & "<br>"
	'response.End 

	set objMessage = CreateObject("CDO.Message")
	'response.write to_email & "<br>"
	'response.write from_email & "<br>"
	'response.write subject & "<br>"
	'response.write content & "<br>"
	'response.End
	set objConfig = createobject("cdo.configuration")

	Schemas = "Http://Schemas.Microsoft.Com/Cdo/Configuration"

	Set Flds = objConfig.Fields
	With Flds
	   Flds.Item(Schemas&"/Sendusing") = 1
	   Flds.Item(Schemas&"/Smtpserver") = "localhost"
	   'Flds.Item(Schemas&"/Smtpserver") = "221.143.41.156"
	   Flds.Item(Schemas&"/Smtpserverport") = 25
	   Flds.Item(Schemas&"/Smtpconnectiontimeout") = 30
	   Flds.Update
	 End With

	Set objMessage.Configuration = objConfig
	'########## 메일 설정 끝


	objMessage.To = to_email	   	    ' 받는사람
	objMessage.From = from_email	   	    ' 보내는사람
	objMessage.bcc = " "	   	    '참조 메일

	objMessage.Subject = subject	   ' 제목
	objMessage.HTMLBody = content	   	   ' 내용

	objMessage.BodyPart.Charset="ks_c_5601-1987"	   	   ' 내용 한글
	objMessage.HTMLBodyPart.Charset="ks_c_5601-1987"
	objMessage.fields.update
	
	if to_email <> "" then
		objMessage.Send	   ' 실제로 보낸는 메소드
	end if

	set objMessage = nothing
	set objConfig = nothing
	set Flds = Nothing

End Function
%>