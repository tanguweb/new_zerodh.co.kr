<%
	Function fn_CodeInfo (p_cd_gubun___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT													"
		f_SQL_____ = f_SQL_____ & "	CD_GUBUN,CD_CODE,NM_CODE,NM_CODE1,NM_CODE2,NM_CODE3		"	'5
		f_SQL_____ = f_SQL_____ & "	,NM_CODE4,NM_CODE5,NM_CODE6,NM_CODE7,NO_ORDER			"	'10
		f_SQL_____ = f_SQL_____ & "	,YN_USE													"
		f_SQL_____ = f_SQL_____ & "FROM														"
		f_SQL_____ = f_SQL_____ & "	T_CODE													"
		f_SQL_____ = f_SQL_____ & "WHERE													"
		f_SQL_____ = f_SQL_____ & "	CD_GUBUN = '" & p_cd_gubun___ & "' AND YN_USE = 'Y'		"
		f_SQL_____ = f_SQL_____ & "ORDER BY													"
		f_SQL_____ = f_SQL_____ & "	NO_ORDER ASC											"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CodeInfo = f_AryList_____
	End Function

	Function fn_CodeInfoWhere (p_cd_gubun___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT													"
		f_SQL_____ = f_SQL_____ & "	CD_GUBUN,CD_CODE,NM_CODE,NM_CODE1,NM_CODE2,NM_CODE3		"	'5
		f_SQL_____ = f_SQL_____ & "	,NM_CODE4,NM_CODE5,NM_CODE6,NM_CODE7,NO_ORDER			"	'10
		f_SQL_____ = f_SQL_____ & "	,YN_USE													"
		f_SQL_____ = f_SQL_____ & "FROM														"
		f_SQL_____ = f_SQL_____ & "	T_CODE													"
		f_SQL_____ = f_SQL_____ & "WHERE													"
		f_SQL_____ = f_SQL_____ & "	CD_GUBUN = '" & p_cd_gubun___ & "' AND YN_USE = 'Y'		"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "ORDER BY													"
		f_SQL_____ = f_SQL_____ & "	NO_ORDER ASC											"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CodeInfoWhere = f_AryList_____
	End Function
	

	Function fn_CodeForBoardCategory (p_cd_gubun___, p_cd_boardid___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT "
		f_SQL_____ = f_SQL_____ & "	A.CD_GUBUN "
		f_SQL_____ = f_SQL_____ & "	, A.CD_CODE "
		f_SQL_____ = f_SQL_____ & "	, A.NM_CODE "
		f_SQL_____ = f_SQL_____ & "	, B.CD_BOARDID "
		f_SQL_____ = f_SQL_____ & "FROM T_CODE A "
		f_SQL_____ = f_SQL_____ & "	LEFT OUTER JOIN T_BOARD_SUB_CATE B "
		f_SQL_____ = f_SQL_____ & "	ON B.CD_SUB_CATE_GUBUN = A.CD_GUBUN "
		f_SQL_____ = f_SQL_____ & "		AND B.CD_SUB_CATE_ID = A.CD_CODE "
		f_SQL_____ = f_SQL_____ & "		AND B.CD_BOARDID = " & p_cd_boardid___ & " "
		f_SQL_____ = f_SQL_____ & "WHERE A.CD_GUBUN = '" & p_cd_gubun___ & "' "
		f_SQL_____ = f_SQL_____ & "ORDER BY A.NO_ORDER ASC "

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CodeForBoardCategory = f_AryList_____
	End Function





		


%>