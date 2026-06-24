<%
'response.write "<!--"
'response.write "SERVER_PROTOCOL [" & Request.ServerVariables("SERVER_PROTOCOL") & "]<br />"
'response.write "SERVER_NAME: [" & Request.ServerVariables("SERVER_NAME") & "]<br />"
'response.write "HTTP_REFERER [" & Request.ServerVariables("HTTP_REFERER") & "]<br />"
'response.write "URL [" & Request.ServerVariables("URL") & "]<br />"
'response.write "QUERY_STRING [" & Request.ServerVariables("QUERY_STRING") & "]<br />"
'response.write "HTTPS [" & Request.ServerVariables("HTTPS") & "]<br />"
'response.write "-->"

'response.write "<!-- 리다이렉트 체크 HTTP_REFERER [" & Request.ServerVariables("HTTP_REFERER") & "] -->"

'Dim p___domain : p___domain = "www.peters.co.kr" ' /include/config.asp 에 정의

If Request.ServerVariables("SERVER_NAME") <> p___domain Then
	
	Dim p___redirect_uri : p___redirect_uri = ""

	If Request.ServerVariables("HTTPS") = "off" Then
		'http
		p___redirect_uri = "http://" & p___domain
	Else
		'https
		p___redirect_uri = "https://" & p___domain
	End If

	If Request.ServerVariables("URL") <> "" Then
		p___redirect_uri = p___redirect_uri & Request.ServerVariables("URL")
	End If

	If Request.ServerVariables("QUERY_STRING") <> "" Then
		p___redirect_uri = p___redirect_uri & "?" & Request.ServerVariables("QUERY_STRING")
	End If

	'response.write "p___redirect_uri: [" & p___redirect_uri & "]<br />"
	
	response.redirect p___redirect_uri
Else

	If Request.ServerVariables("HTTPS") = "off" Then
		'http
		p___redirect_uri = "https://" & p___domain

		If Request.ServerVariables("URL") <> "" Then
			p___redirect_uri = p___redirect_uri & Request.ServerVariables("URL")
		End If

		If Request.ServerVariables("QUERY_STRING") <> "" Then
			p___redirect_uri = p___redirect_uri & "?" & Request.ServerVariables("QUERY_STRING")
		End If

		'response.write "p___redirect_uri: [" & p___redirect_uri & "]<br />"
		
		response.redirect p___redirect_uri
	End If


End If 
%>