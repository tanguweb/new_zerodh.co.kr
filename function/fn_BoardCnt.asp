<%
	' config.asp 인클루드 아래에 추가합니다.
	Sub fn_BoardCnt (f_CD_BOARDID___)

		Dim f_objADO_____, f_Rs_____, f_SQL_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = f_SQL_____ & " UPDATE T_BOARDDETAIL SET		"
		f_SQL_____ = f_SQL_____ & "	NO_VIEWCNT = NO_VIEWCNT + 1		"
		f_SQL_____ = f_SQL_____ & "	WHERE CD_BOARDID = " & f_CD_BOARDID___ & "  "

		f_objADO_____.setSQL(f_SQL_____)
		f_objADO_____.getExec()

		Set f_objADO_____ = Nothing

	End Sub

%>