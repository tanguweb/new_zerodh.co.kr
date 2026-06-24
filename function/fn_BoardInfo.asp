<%
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' [주의사항]
	' 1. 해당 파일 사용시 config.asp파일이 인클루드 되어있어야 하고, 해당파일 아래에 인클루드 시켜준다.
	' 2. 해당 함수의 SELECT 절의 컬럼위치를 변경하지 않는다.
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
	' 게시판리스트
	Function fn_BoardInfo (p_cd_boardcd___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardInfo = f_AryList_____
	End Function

	' 게시판리스트(페이지포함)
	Function fn_BoardInfo_paging (p_cd_boardcd___, p_page___, p_pgsize___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "DECLARE @CNT int																					"
		f_SQL_____ = f_SQL_____ & "DECLARE @SNUM int																				"
		f_SQL_____ = f_SQL_____ & "DECLARE @ENUM int																				"
		f_SQL_____ = f_SQL_____ & "SET @ENUM = " & p_page___ & " * " & p_pgsize___ & "												"
		f_SQL_____ = f_SQL_____ & "SET @SNUM = (@ENUM - " & p_pgsize___ & ")+1														"
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT = COUNT(*)																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT AS CNT, *																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	(																								"
		f_SQL_____ = f_SQL_____ & "		SELECT																						"
		f_SQL_____ = f_SQL_____ & "			ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]	"	'1
		f_SQL_____ = f_SQL_____ & "			,BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD							"	'5
		f_SQL_____ = f_SQL_____ & "			,BD.YN_DISPLAY,BD.NO_VIEWCNT,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE						"	'10
		f_SQL_____ = f_SQL_____ & "			,BD.NM_CONTENTS,BD.NM_FIELD_1,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4					"	'15
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_5,BD.NM_FIELD_6,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9					"	'20
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_10,BD.CD_INUSER,BD.DT_INSYSDATE											"	'23
		f_SQL_____ = f_SQL_____ & "			,ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC) AS [NoDesc]	"	'24
		f_SQL_____ = f_SQL_____ & "			,BD.CD_BOARDKEY																			"	'25
		f_SQL_____ = f_SQL_____ & "		FROM																						"
		f_SQL_____ = f_SQL_____ & "			T_BOARDDETAIL AS BD																		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_BOARDMAST AS BM														"
		f_SQL_____ = f_SQL_____ & "				ON BM.CD_BOARDCD = BD.CD_BOARDCD													"
		f_SQL_____ = f_SQL_____ & "		WHERE																						"
		f_SQL_____ = f_SQL_____ & "			BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'													"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'											"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "	) AS J																							"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM																		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	NO ASC																							"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardInfo_paging = f_AryList_____
	End Function
	
	' 게시판리스트(페이지포함)
	Function fn_BoardInfo_paging_join_member (p_cd_boardcd___, p_page___, p_pgsize___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "DECLARE @CNT int																					"
		f_SQL_____ = f_SQL_____ & "DECLARE @SNUM int																				"
		f_SQL_____ = f_SQL_____ & "DECLARE @ENUM int																				"
		f_SQL_____ = f_SQL_____ & "SET @ENUM = " & p_page___ & " * " & p_pgsize___ & "												"
		f_SQL_____ = f_SQL_____ & "SET @SNUM = (@ENUM - " & p_pgsize___ & ")+1														"
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT = COUNT(*)																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN DB_MEMBER.dbo.T_MEMBER AS MB														"
		f_SQL_____ = f_SQL_____ & "		ON MB.CD_PRIMARYID = BD.CD_INUSER															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT AS CNT, *																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	(																								"
		f_SQL_____ = f_SQL_____ & "		SELECT																						"
		f_SQL_____ = f_SQL_____ & "			ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]	"	'1
		f_SQL_____ = f_SQL_____ & "			,BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD							"	'5
		f_SQL_____ = f_SQL_____ & "			,BD.YN_DISPLAY,BD.NO_VIEWCNT,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE						"	'10
		f_SQL_____ = f_SQL_____ & "			,BD.NM_CONTENTS,BD.NM_FIELD_1,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4					"	'15
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_5,BD.NM_FIELD_6,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9					"	'20
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_10,BD.CD_INUSER,BD.DT_INSYSDATE											"	'23
		f_SQL_____ = f_SQL_____ & "			,ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC) AS [NoDesc]	"	'24
		f_SQL_____ = f_SQL_____ & "			,BD.CD_BOARDKEY, DB_INTHEF.dbo.UFN_GET_AESDEC(MB.NM_USERNM)	AS NM_USERNM, MB.CD_USERID, MB.YN_ADMIN		"	'28
		f_SQL_____ = f_SQL_____ & "		FROM																						"
		f_SQL_____ = f_SQL_____ & "			T_BOARDDETAIL AS BD																		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_BOARDMAST AS BM														"
		f_SQL_____ = f_SQL_____ & "				ON BM.CD_BOARDCD = BD.CD_BOARDCD													"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN DB_MEMBER.dbo.T_MEMBER AS MB												"
		f_SQL_____ = f_SQL_____ & "				ON MB.CD_PRIMARYID = BD.CD_INUSER													"
		f_SQL_____ = f_SQL_____ & "		WHERE																						"
		f_SQL_____ = f_SQL_____ & "			BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'													"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'											"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "	) AS J																							"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM																		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	NO ASC																							"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardInfo_paging_join_member = f_AryList_____
	End Function

	' 게시글보기
	Function fn_BoardView (f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDID = '" & f_CD_BOARDID___ & "'													"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardView = f_AryList_____
	End Function

	' 이전글보기
	Function fn_BoardBefore (f_CD_BOARDCD___, f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & f_CD_BOARDCD___ & "' AND BD.CD_BOARDID < '" & f_CD_BOARDID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardBefore = f_AryList_____
	End Function

	' 다음글보기
	Function fn_BoardAfter (f_CD_BOARDCD___, f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & f_CD_BOARDCD___ & "' AND BD.CD_BOARDID > '" & f_CD_BOARDID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC									"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardAfter = f_AryList_____
	End Function
	
	' Q&A리스트
	Function fn_CustomerInfo (p_cd_boardcd___, p_page___, p_pgsize___, p_UserID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "DECLARE @CNT int																					"
		f_SQL_____ = f_SQL_____ & "DECLARE @SNUM int																				"
		f_SQL_____ = f_SQL_____ & "DECLARE @ENUM int																				"
		f_SQL_____ = f_SQL_____ & "SET @ENUM = " & p_page___ & " * " & p_pgsize___ & "												"
		f_SQL_____ = f_SQL_____ & "SET @SNUM = (@ENUM - " & p_pgsize___ & ")+1														"
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT = COUNT(*)																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_CODE AS CD																		"
		f_SQL_____ = f_SQL_____ & "		ON CD.CD_CODE = BD.NM_FIELD_1 AND CD.CD_GUBUN = 'CUSTOMER'									"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDKEY IN (SELECT CD_BOARDKEY FROM T_BOARDDETAIL WHERE CD_BOARDCD = '" & p_cd_boardcd___ & "' AND CD_INUSER = '" & p_UserID___ & "' AND NO_BOARD_DEPTH = 1) "
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT AS CNT, *																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	(																								"
		f_SQL_____ = f_SQL_____ & "		SELECT																						"
		f_SQL_____ = f_SQL_____ & "			ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]	"	'1
		f_SQL_____ = f_SQL_____ & "			,BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD							"	'5
		f_SQL_____ = f_SQL_____ & "			,BD.YN_DISPLAY,BD.NO_VIEWCNT,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE						"	'10
		f_SQL_____ = f_SQL_____ & "			,BD.NM_CONTENTS,BD.NM_FIELD_1,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4					"	'15
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_5,BD.NM_FIELD_6,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9					"	'20
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_10,BD.CD_INUSER,BD.DT_INSYSDATE											"	'23
		f_SQL_____ = f_SQL_____ & "			,ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDID ASC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC) AS [NoDesc]	"	'24
		f_SQL_____ = f_SQL_____ & "			,CD.NM_CODE																				"	'25
		f_SQL_____ = f_SQL_____ & "		FROM																						"
		f_SQL_____ = f_SQL_____ & "			T_BOARDDETAIL AS BD																		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_BOARDMAST AS BM														"
		f_SQL_____ = f_SQL_____ & "				ON BM.CD_BOARDCD = BD.CD_BOARDCD													"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS CD																		"
		f_SQL_____ = f_SQL_____ & "				ON CD.CD_CODE = BD.NM_FIELD_1 AND CD.CD_GUBUN = 'CUSTOMER'									"
		f_SQL_____ = f_SQL_____ & "		WHERE																						"
		f_SQL_____ = f_SQL_____ & "			BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'													"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'											"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDKEY IN (SELECT CD_BOARDKEY FROM T_BOARDDETAIL WHERE CD_BOARDCD = '" & p_cd_boardcd___ & "' AND CD_INUSER = '" & p_UserID___ & "' AND NO_BOARD_DEPTH = 1) "
		f_SQL_____ = f_SQL_____ & "	) AS J																							"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM																		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	NO ASC																							"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerInfo = f_AryList_____
	End Function

	' Q&A리스트
	Function fn_CustomerInfoJoinMember (p_cd_boardcd___, p_page___, p_pgsize___, p_UserID___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "DECLARE @CNT int																					"
		f_SQL_____ = f_SQL_____ & "DECLARE @SNUM int																				"
		f_SQL_____ = f_SQL_____ & "DECLARE @ENUM int																				"
		f_SQL_____ = f_SQL_____ & "SET @ENUM = " & p_page___ & " * " & p_pgsize___ & "												"
		f_SQL_____ = f_SQL_____ & "SET @SNUM = (@ENUM - " & p_pgsize___ & ")+1														"
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT = COUNT(*)																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_CODE AS CD																		"
		f_SQL_____ = f_SQL_____ & "		ON CD.CD_CODE = BD.NM_FIELD_1 AND CD.CD_GUBUN = 'CUSTOMER'									"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN DB_MEMBER.dbo.T_MEMBER AS MB														"
		f_SQL_____ = f_SQL_____ & "		ON MB.CD_PRIMARYID = BD.CD_INUSER															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDKEY IN (SELECT CD_BOARDKEY FROM T_BOARDDETAIL WHERE CD_BOARDCD = '" & p_cd_boardcd___ & "' AND CD_INUSER = '" & p_UserID___ & "' AND NO_BOARD_DEPTH = 1) "
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT AS CNT, *																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	(																								"
		f_SQL_____ = f_SQL_____ & "		SELECT																						"
		f_SQL_____ = f_SQL_____ & "			ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]	"	'1
		f_SQL_____ = f_SQL_____ & "			,BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD							"	'5
		f_SQL_____ = f_SQL_____ & "			,BD.YN_DISPLAY,BD.NO_VIEWCNT,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE						"	'10
		f_SQL_____ = f_SQL_____ & "			,BD.NM_CONTENTS,BD.NM_FIELD_1,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4					"	'15
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_5,BD.NM_FIELD_6,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9					"	'20
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_10,BD.CD_INUSER,BD.DT_INSYSDATE											"	'23
		f_SQL_____ = f_SQL_____ & "			,ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDID ASC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC) AS [NoDesc]	"	'24
		f_SQL_____ = f_SQL_____ & "			,CD.NM_CODE, DB_INTHEF.dbo.UFN_GET_AESDEC(MB.NM_USERNM)	AS NM_USERNM, MB.CD_USERID, MB.YN_ADMIN									"	'28
		f_SQL_____ = f_SQL_____ & "		FROM																						"
		f_SQL_____ = f_SQL_____ & "			T_BOARDDETAIL AS BD																		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_BOARDMAST AS BM														"
		f_SQL_____ = f_SQL_____ & "				ON BM.CD_BOARDCD = BD.CD_BOARDCD													"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS CD																"
		f_SQL_____ = f_SQL_____ & "				ON CD.CD_CODE = BD.NM_FIELD_1 AND CD.CD_GUBUN = 'CUSTOMER'							"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN DB_MEMBER.dbo.T_MEMBER AS MB												"
		f_SQL_____ = f_SQL_____ & "				ON MB.CD_PRIMARYID = BD.CD_INUSER													"
		f_SQL_____ = f_SQL_____ & "		WHERE																						"
		f_SQL_____ = f_SQL_____ & "			BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'													"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'											"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDKEY IN (SELECT CD_BOARDKEY FROM T_BOARDDETAIL WHERE CD_BOARDCD = '" & p_cd_boardcd___ & "' AND CD_INUSER = '" & p_UserID___ & "' AND NO_BOARD_DEPTH = 1) "
		f_SQL_____ = f_SQL_____ & "	) AS J																							"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM																		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	NO ASC																							"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerInfoJoinMember = f_AryList_____
	End Function

	' Q&A보기
	Function fn_CustomerView (f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE,CD.NM_CODE																		"	'22
		f_SQL_____ = f_SQL_____ & " ,BD.CD_BOARDKEY,BM.NM_BOARDNM																	"	'24
		f_SQL_____ = f_SQL_____ & ",(																													"
		f_SQL_____ = f_SQL_____ & "		SELECT CASE WHEN COUNT(*) > 0 THEN 'Y' ELSE 'N' END																"
		f_SQL_____ = f_SQL_____ & "		FROM T_BOARDDETAIL 																								"
		f_SQL_____ = f_SQL_____ & "		WHERE YN_USE = 'Y' AND YN_DISPLAY = 'Y' 																		"
		f_SQL_____ = f_SQL_____ & "			AND CD_BOARDKEY = (SELECT CD_BOARDKEY FROM T_BOARDDETAIL WHERE CD_BOARDID = '" & f_CD_BOARDID___ & "') 		"
		f_SQL_____ = f_SQL_____ & "			AND NO_BOARD_SEQ > 1																						"
		f_SQL_____ = f_SQL_____ & ")																													" '23
		f_SQL_____ = f_SQL_____ & ",CD.NM_CODE, DB_INTHEF.dbo.UFN_GET_AESDEC(MB.NM_USERNM)	AS NM_USERNM, MB.CD_USERID, MB.YN_ADMIN						"	'29
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_CODE AS CD																		"
		f_SQL_____ = f_SQL_____ & "		ON CD.CD_CODE = BD.NM_FIELD_1 AND CD.CD_GUBUN = 'CUSTOMER'									"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN DB_MEMBER.dbo.T_MEMBER AS MB												"
		f_SQL_____ = f_SQL_____ & "		ON MB.CD_PRIMARYID = BD.CD_INUSER													"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDID = '" & f_CD_BOARDID___ & "'													"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerView = f_AryList_____
	End Function

	' 게시판리스트 WHERE 절 추가 가능 쿼리
	Function fn_BoardInfoWhere (p_cd_boardcd___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardInfoWhere = f_AryList_____
	End Function

	' 게시판리스트(시작일:NM_FIELD_1,종료일 검사:NM_FIELD_2)
	Function fn_BoardInfoDate (p_cd_boardcd___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & "	AND (CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8),GETDATE(),108),':','') >= NM_FIELD_1) "
		f_SQL_____ = f_SQL_____ & "	AND (CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8),GETDATE(),108),':','') <= NM_FIELD_2) "
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardInfoDate = f_AryList_____
	End Function

	' 게시글보기 Where 절 추가
	Function fn_BoardViewWhere (f_CD_BOARDID___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDID = '" & f_CD_BOARDID___ & "'													"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardViewWhere = f_AryList_____
	End Function

	' Q&A리스트
	Function fn_CustomerInfoWhere (p_cd_boardcd___, p_page___, p_pgsize___, p_UserID___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "DECLARE @CNT int																					"
		f_SQL_____ = f_SQL_____ & "DECLARE @SNUM int																				"
		f_SQL_____ = f_SQL_____ & "DECLARE @ENUM int																				"
		f_SQL_____ = f_SQL_____ & "SET @ENUM = " & p_page___ & " * " & p_pgsize___ & "												"
		f_SQL_____ = f_SQL_____ & "SET @SNUM = (@ENUM - " & p_pgsize___ & ")+1														"
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT = COUNT(*)																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_CODE AS CD																		"
		f_SQL_____ = f_SQL_____ & "		ON CD.CD_CODE = BD.NM_FIELD_1 AND CD.CD_GUBUN = 'CUSTOMER'									"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'													"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDKEY IN (SELECT CD_BOARDKEY FROM T_BOARDDETAIL WHERE CD_BOARDCD = '" & p_cd_boardcd___ & "' AND CD_INUSER = '" & p_UserID___ & "' AND NO_BOARD_DEPTH = 1) "
		f_SQL_____ = f_SQL_____ & "																									"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT AS CNT, *																			"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	(																								"
		f_SQL_____ = f_SQL_____ & "		SELECT																						"
		f_SQL_____ = f_SQL_____ & "			ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]	"	'1
		f_SQL_____ = f_SQL_____ & "			,BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD							"	'5
		f_SQL_____ = f_SQL_____ & "			,BD.YN_DISPLAY,BD.NO_VIEWCNT,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE						"	'10
		f_SQL_____ = f_SQL_____ & "			,BD.NM_CONTENTS,BD.NM_FIELD_1,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4					"	'15
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_5,BD.NM_FIELD_6,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9					"	'20
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_10,BD.CD_INUSER,BD.DT_INSYSDATE											"	'23
		f_SQL_____ = f_SQL_____ & "			,ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDID ASC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC) AS [NoDesc]	"	'24
		f_SQL_____ = f_SQL_____ & "			,CD.NM_CODE																				"	'25
		f_SQL_____ = f_SQL_____ & "		FROM																						"
		f_SQL_____ = f_SQL_____ & "			T_BOARDDETAIL AS BD																		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_BOARDMAST AS BM														"
		f_SQL_____ = f_SQL_____ & "				ON BM.CD_BOARDCD = BD.CD_BOARDCD													"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS CD																		"
		f_SQL_____ = f_SQL_____ & "				ON CD.CD_CODE = BD.NM_FIELD_1 AND CD.CD_GUBUN = 'CUSTOMER'									"
		f_SQL_____ = f_SQL_____ & "		WHERE																						"
		f_SQL_____ = f_SQL_____ & "			BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'													"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'											"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDKEY IN (SELECT CD_BOARDKEY FROM T_BOARDDETAIL WHERE CD_BOARDCD = '" & p_cd_boardcd___ & "' AND CD_INUSER = '" & p_UserID___ & "' AND NO_BOARD_DEPTH = 1) "
		f_SQL_____ = f_SQL_____ & "	) AS J																							"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM																		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	NO ASC																							"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerInfoWhere = f_AryList_____
	End Function

	Function fn_PopupInfo (p_CD_BRANDID)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '1056'																		"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_1 = '"&p_CD_BRANDID&"'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_2 <= CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8),GETDATE(),108),':','')"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_3 >= CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8),GETDATE(),108),':','')"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_PopupInfo = f_AryList_____
	End Function

	' 이전글보기
	Function fn_CustomerBefore (f_CD_BOARDCD___, f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & f_CD_BOARDCD___ & "' AND BD.CD_BOARDID < '" & f_CD_BOARDID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID DESC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ DESC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerBefore = f_AryList_____
	End Function

	' 다음글보기
	Function fn_CustomerAfter (f_CD_BOARDCD___, f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & f_CD_BOARDCD___ & "' AND BD.CD_BOARDID > '" & f_CD_BOARDID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC									"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerAfter = f_AryList_____
	End Function

	' 이전글보기(통합 고객상담 어드민용)
	Function fn_BoardBeforeMerge (f_CD_BOARDTYPE___, f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BM.CD_BOARDTYPE = '" & f_CD_BOARDTYPE___ & "' AND BD.CD_BOARDID < '" & f_CD_BOARDID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardBeforeMerge = f_AryList_____
	End Function

	' 다음글보기(통합 고객상담 어드민용)
	Function fn_BoardAfterMerge (f_CD_BOARDTYPE___, f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BM.CD_BOARDTYPE = '" & f_CD_BOARDTYPE___ & "' AND BD.CD_BOARDID > '" & f_CD_BOARDID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC									"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardAfterMerge = f_AryList_____
	End Function

	'카다로그 신청 팝업
	Function fn_ApplyCatalogue (p_CD_BOARDCD)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '"&p_CD_BOARDCD&"'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.NO_BOARD_DEPTH = '1' AND BD.NO_BOARD_SEQ = '1'											"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_3 <= CONVERT(VARCHAR(8),GETDATE(),112)											"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_4 >= CONVERT(VARCHAR(8),GETDATE(),112)											"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_ApplyCatalogue = f_AryList_____
	End Function

	' 이전글보기(각 브랜드에 사용)
	Function fn_CustomerBefore_v2 (f_CD_BOARDCD___, f_CD_BOARDID___, f_m_PrimaryID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & f_CD_BOARDCD___ & "' AND BD.CD_BOARDID < '" & f_CD_BOARDID___ & "' AND BD.CD_INUSER = '" & f_m_PrimaryID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID DESC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ DESC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerBefore_v2 = f_AryList_____
	End Function

	' 다음글보기(각 브랜드에 사용)
	Function fn_CustomerAfter_v2 (f_CD_BOARDCD___, f_CD_BOARDID___, f_m_PrimaryID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT TOP 1																						"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '" & f_CD_BOARDCD___ & "' AND BD.CD_BOARDID > '" & f_CD_BOARDID___ & "' AND BD.CD_INUSER = '" & f_m_PrimaryID___ & "' AND BD.NO_BOARD_DEPTH = 1 AND BD.NO_BOARD_SEQ = 1		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC									"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_CustomerAfter_v2 = f_AryList_____
	End Function

	'공모전 신청 팝업 (사용 site : S+ByTrugen)
	Function fn_ApplyContest (p_CD_BOARDCD)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '"&p_CD_BOARDCD&"'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.NO_BOARD_DEPTH = '1' AND BD.NO_BOARD_SEQ = '1'											"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_2 <= CONVERT(VARCHAR(8),GETDATE(),112)											"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_3 >= CONVERT(VARCHAR(8),GETDATE(),112)											"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_ApplyContest = f_AryList_____
	End Function


	'프레디 온라인 프로모션 팝업 (사용 site : www.freddykorea.co.kr)
	Function fn_Promotion (p_CD_BOARDCD)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO

		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "	SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDCD = '"&p_CD_BOARDCD&"'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.NO_BOARD_DEPTH = '1' AND BD.NO_BOARD_SEQ = '1'											"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_5 <= CONVERT(VARCHAR(8),GETDATE(),112)											"
		f_SQL_____ = f_SQL_____ & "	AND BD.NM_FIELD_6 >= CONVERT(VARCHAR(8),GETDATE(),112)											"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ Asc				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_Promotion = f_AryList_____
	End Function 
%>


