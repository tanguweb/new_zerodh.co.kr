<%
	Function fn_PopupInfo ()
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																			"
		f_SQL_____ = f_SQL_____ & "		CD_BOARDID, CD_BOARDCD, NM_TITLE, NM_CONTENTS								"
		f_SQL_____ = f_SQL_____ & "		, NM_FIELD_1, NM_FIELD_2, NM_FIELD_3, NM_FIELD_4, NM_FIELD_5				"
		f_SQL_____ = f_SQL_____ & "		, NM_FIELD_6, NM_FIELD_7, NM_FIELD_8, NM_FIELD_9, NM_FIELD_10				"
		f_SQL_____ = f_SQL_____ & "		, NM_FIELD_11, NM_FIELD_12, NM_FIELD_13, NM_FIELD_14, NM_FIELD_15			"
		f_SQL_____ = f_SQL_____ & "FROM																				"
		f_SQL_____ = f_SQL_____ & "		T_BOARDDETAIL AS BD															"
		f_SQL_____ = f_SQL_____ & "WHERE CD_BOARDCD = '1001' AND YN_USE = 'Y' AND YN_DISPLAY = 'Y'					"
		f_SQL_____ = f_SQL_____ & "		AND NM_FIELD_1 <= CONVERT(VARCHAR, GETDATE(), 112)							"
		f_SQL_____ = f_SQL_____ & "		AND NM_FIELD_2 >= CONVERT(VARCHAR, GETDATE(), 112)							"
		f_SQL_____ = f_SQL_____ & "ORDER BY																			"
		f_SQL_____ = f_SQL_____ & "		CD_BOARDID DESC																"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_PopupInfo = f_AryList_____
	End Function
%>