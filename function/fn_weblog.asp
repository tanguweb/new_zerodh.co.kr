<%
	Function ff_GetValue_____(val)
		Dim rt : rt = "null"
		Dim va___
		For Each va___ In Request.ServerVariables
			If va___ = val Then
				rt = Request.ServerVariables(val)
			End If
		Next
		ff_GetValue_____ = rt
	End Function

	Function weblog_pageconnect()

		On Error Resume Next

		Dim f_Par_____ : f_Par_____ = Split(ff_GetValue_____("QUERY_STRING"),"=")
		Dim f_adoCmd_____
		SET f_adoCmd_____	=	Server.CreateObject("ADODB.Command")
			with f_adoCmd_____
			.ActiveConnection	= m_DB_Log
			.CommandType		= adCmdStoredProc
			.CommandText		= "USP_WRITELOG"
			.Parameters.Append	.CreateParameter("@CD_TYPE"				, adChar,		adParamInput,	1,		"P")
			.Parameters.Append	.CreateParameter("@NM_SESSIONID"		, adVarchar,	adParamInput,	200,	Session.SessionID)
			.Parameters.Append	.CreateParameter("@NM_REMOTE_ADDR"		, adVarchar,	adParamInput,	60,		ff_GetValue_____("REMOTE_ADDR"))
			.Parameters.Append	.CreateParameter("@NM_QUERY_STRING"		, adVarchar,	adParamInput,	200,	ff_GetValue_____("QUERY_STRING"))
			.Parameters.Append	.CreateParameter("@NM_DOMAIN"			, adVarchar,	adParamInput,	100,	ff_GetValue_____("HTTP_HOST"))
			.Parameters.Append	.CreateParameter("@NM_URL"				, adVarchar,	adParamInput,	200,	ff_GetValue_____("URL"))
			.Parameters.Append	.CreateParameter("@NM_BROWSER"			, adVarchar,	adParamInput,	1000,	ff_GetValue_____("HTTP_USER_AGENT"))
			.Execute
			end with
		SET f_adoCmd_____ = Nothing
	End Function 

	weblog_pageconnect()
%>