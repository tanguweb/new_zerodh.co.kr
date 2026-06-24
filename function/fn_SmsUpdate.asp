<%
	' config.asp 인클루드 아래에 추가합니다.
	Sub fn_SmsUpdate (f_ID___, f_rtn___)

		Dim f_objADO_____, f_Rs_____, f_SQL_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = f_SQL_____ & " UPDATE T_SMS SET					"
		f_SQL_____ = f_SQL_____ & "	NM_RESULT = '" & f_rtn___ & "'		"
		f_SQL_____ = f_SQL_____ & "	WHERE NO_LOG = " & f_ID___ & "		"

		f_objADO_____.setSQL(f_SQL_____)
		f_objADO_____.getExec()

		Set f_objADO_____ = Nothing

	End Sub

%>