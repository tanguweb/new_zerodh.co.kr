<%
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' [주의사항]
	' 1. 해당 파일 사용시 config.asp파일이 인클루드 되어있어야 하고, 해당파일 아래에 인클루드 시켜준다.
	' 2. 해당 함수의 SELECT 절의 컬럼위치를 변경하지 않는다.
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
	' 게시판리스트
	' 게시판리스트(페이지포함)
	Function fn_BoardList_paging (p_cd_boardcd___, p_page___, p_pgsize___, p_AddWhereQuery)
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
		f_SQL_____ = f_SQL_____ & "			AND BD.NO_BOARD_DEPTH = '1' AND BD.NO_BOARD_SEQ = '1'									"
		'f_SQL_____ = f_SQL_____ & "			AND BD.NM_FIELD_2 = 'counsel' 									"
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
		f_SQL_____ = f_SQL_____ & "			,(SELECT COUNT(*) FROM T_BOARDREPLY WHERE CD_BOARDID = BD.CD_BOARDID AND YN_USE = 'Y') AS REPLYCNT	"	'26
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_11,BD.NM_FIELD_12,BD.NM_FIELD_13,BD.NM_FIELD_14							"	'30
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_15																			"	'31
		f_SQL_____ = f_SQL_____ & "			,(SELECT COUNT(*) FROM T_BOARDDETAIL WHERE CD_BOARDKEY = BD.CD_BOARDID AND YN_USE = 'Y' AND NO_BOARD_DEPTH > 1) AS ANSHERCNT	"	'32
		f_SQL_____ = f_SQL_____ & "		FROM																						"
		f_SQL_____ = f_SQL_____ & "			T_BOARDDETAIL AS BD																		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_BOARDMAST AS BM														"
		f_SQL_____ = f_SQL_____ & "				ON BM.CD_BOARDCD = BD.CD_BOARDCD													"
		f_SQL_____ = f_SQL_____ & "		WHERE																						"
		f_SQL_____ = f_SQL_____ & "			BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'													"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'											"
		f_SQL_____ = f_SQL_____ & "			AND BD.NO_BOARD_DEPTH = '1' AND BD.NO_BOARD_SEQ = '1'									"
		'f_SQL_____ = f_SQL_____ & "			AND BD.NM_FIELD_2 = 'counsel' 									"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "	) AS J																							"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM																		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	NO ASC																							"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardList_paging = f_AryList_____
	End Function




     Function fn_BoardList_paging2 (p_cd_boardcd___, p_page___, p_pgsize___, p_AddWhereQuery)
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
		f_SQL_____ = f_SQL_____ & "			AND BD.NO_BOARD_DEPTH = '1' AND BD.NO_BOARD_SEQ = '1'									"
		f_SQL_____ = f_SQL_____ & "			AND BD.NM_FIELD_2 = 'upgrade' 									"
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
		f_SQL_____ = f_SQL_____ & "			,(SELECT COUNT(*) FROM T_BOARDREPLY WHERE CD_BOARDID = BD.CD_BOARDID AND YN_USE = 'Y') AS REPLYCNT	"	'26
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_11,BD.NM_FIELD_12,BD.NM_FIELD_13,BD.NM_FIELD_14							"	'30
		f_SQL_____ = f_SQL_____ & "			,BD.NM_FIELD_15																			"	'31
		f_SQL_____ = f_SQL_____ & "			,(SELECT COUNT(*) FROM T_BOARDDETAIL WHERE CD_BOARDKEY = BD.CD_BOARDID AND YN_USE = 'Y' AND NO_BOARD_DEPTH > 1) AS ANSHERCNT	"	'32
		f_SQL_____ = f_SQL_____ & "		FROM																						"
		f_SQL_____ = f_SQL_____ & "			T_BOARDDETAIL AS BD																		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_BOARDMAST AS BM														"
		f_SQL_____ = f_SQL_____ & "				ON BM.CD_BOARDCD = BD.CD_BOARDCD													"
		f_SQL_____ = f_SQL_____ & "		WHERE																						"
		f_SQL_____ = f_SQL_____ & "			BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'													"
		f_SQL_____ = f_SQL_____ & "			AND BD.CD_BOARDCD = '" & p_cd_boardcd___ & "'											"
		f_SQL_____ = f_SQL_____ & "			AND BD.NO_BOARD_DEPTH = '1' AND BD.NO_BOARD_SEQ = '1'									"
		f_SQL_____ = f_SQL_____ & "			AND BD.NM_FIELD_2 = 'upgrade' 									"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "	) AS J																							"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM																		"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	NO ASC																							"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardList_paging2 = f_AryList_____
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
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"	'21
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_11,BD.NM_FIELD_12,BD.NM_FIELD_13,BD.NM_FIELD_14									"	'25
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_15																					"	'26
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

	Function fn_BoardView_addQuery (f_CD_BOARDID___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"	'21
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_11,BD.NM_FIELD_12,BD.NM_FIELD_13,BD.NM_FIELD_14									"	'25
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_15																					"	'26
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

		fn_BoardView_addQuery = f_AryList_____
	End Function

	Function fn_BoardAnswer (f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	BD.CD_BOARDID,BD.NO_BOARD_DEPTH,BD.NO_BOARD_SEQ,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT		"	'5
		f_SQL_____ = f_SQL_____ & "	,BD.NO_VIEWORDER,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS,BD.NM_FIELD_1								"	'10
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5,BD.NM_FIELD_6							"	'15
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10,BD.CD_INUSER							"	'20
		f_SQL_____ = f_SQL_____ & "	,BD.DT_INSYSDATE																				"	'21
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_11,BD.NM_FIELD_12,BD.NM_FIELD_13,BD.NM_FIELD_14									"	'25
		f_SQL_____ = f_SQL_____ & "	,BD.NM_FIELD_15																					"	'26
		f_SQL_____ = f_SQL_____ & "FROM																								"
		f_SQL_____ = f_SQL_____ & "	T_BOARDDETAIL AS BD																				"
		f_SQL_____ = f_SQL_____ & "		INNER JOIN T_BOARDMAST AS BM																"
		f_SQL_____ = f_SQL_____ & "		ON BM.CD_BOARDCD = BD.CD_BOARDCD															"
		f_SQL_____ = f_SQL_____ & "WHERE																							"
		f_SQL_____ = f_SQL_____ & "	BD.YN_USE = 'Y' AND BD.YN_DISPLAY = 'Y'															"
		f_SQL_____ = f_SQL_____ & "	AND BD.CD_BOARDKEY = '" & f_CD_BOARDID___ & "' AND BD.NO_BOARD_DEPTH > 1						"
		f_SQL_____ = f_SQL_____ & "ORDER BY																							"
		f_SQL_____ = f_SQL_____ & "	BD.NO_VIEWORDER ASC, BD.CD_BOARDID DESC, BD.NO_BOARD_DEPTH ASC, BD.NO_BOARD_SEQ ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardAnswer = f_AryList_____
	End Function

	' 게시글의 덧글보기
	Function fn_BoardReply (f_CD_BOARDID___)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
			
		f_SQL_____ = ""
		f_SQL_____ = f_SQL_____ & "SELECT																							"
		f_SQL_____ = f_SQL_____ & "	CD_REPLYID, CD_BOARDID, NM_REPLY, YN_USE, CD_INUSER, NO_INIPADDR								"	'5
		f_SQL_____ = f_SQL_____ & "	, DT_INSYSDATE, CD_MDUSER, NO_MDIPADDR, DT_MDSYSDATE, NM_INUSER									"	'10
		f_SQL_____ = f_SQL_____ & "	, NM_PASSWORD																					"	'11
		f_SQL_____ = f_SQL_____ & "FROM T_BOARDREPLY 																				"
		f_SQL_____ = f_SQL_____ & "WHERE CD_BOARDID = '" & f_CD_BOARDID___ & "' AND YN_USE = 'Y'									"
		f_SQL_____ = f_SQL_____ & "ORDER BY	CD_REPLYID DESC																			"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_BoardReply = f_AryList_____
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
	
%>


