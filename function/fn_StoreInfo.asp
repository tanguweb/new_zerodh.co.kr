<%
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' [주의사항]
	' 1. 해당 파일 사용시 config.asp파일이 인클루드 되어있어야 하고, 해당파일 아래에 인클루드 시켜준다.
	' 2. 해당 함수의 SELECT 절의 컬럼위치를 변경하지 않는다.
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
	Function fn_StoreList_paging (p_cd_brandid___, p_page___, p_pgsize___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
		
		f_SQL_____ = "DECLARE @CNT int									"
		f_SQL_____ = f_SQL_____ & "DECLARE @SNUM int							"
		f_SQL_____ = f_SQL_____ & "DECLARE @ENUM int							"
		f_SQL_____ = f_SQL_____ & "SET @ENUM = " & p_page___ & " * " & p_pgsize___ & "	"
		f_SQL_____ = f_SQL_____ & "SET @SNUM = (@ENUM - " & p_pgsize___ & ")+1		"
		f_SQL_____ = f_SQL_____ & "											"
		f_SQL_____ = f_SQL_____ & "SELECT @CNT = COUNT(*)						"
		f_SQL_____ = f_SQL_____ & "		FROM																				"
		f_SQL_____ = f_SQL_____ & "			T_STORE AS A																	"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS B ON B.CD_GUBUN = 'BRANDID' AND A.CD_BRANDID = B.CD_CODE		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS C ON C.CD_GUBUN = 'STORETYPE' AND A.NM_STORETYPE = C.CD_CODE	"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS D ON D.CD_GUBUN = 'AREA' AND A.NM_AREA = D.CD_CODE				"	
		f_SQL_____ = f_SQL_____ & "		WHERE A.YN_USE = 'Y' AND C.YN_USE = 'Y' AND A.YN_ING = 'Y' AND B.CD_CODE = '" & p_cd_brandid___ & "'"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery	
		f_SQL_____ = f_SQL_____ & ""
		f_SQL_____ = f_SQL_____ & "SELECT																						"
		f_SQL_____ = f_SQL_____ & "	@CNT AS CNT, NO, CD_STOREID, CD_BRANDID, CD_STORECD, NM_STORETYPE, NO_VIEWORDER, NM_AREA, NM_STORE, NM_ADDR, NO_TEL1, NO_ORDER "
										'0		'1		'2			'3			'4			'5				'6			'7		'8			'9		'10		'11
		f_SQL_____ = f_SQL_____ & "FROM																						"
		f_SQL_____ = f_SQL_____ & "	(																						"
		f_SQL_____ = f_SQL_____ & "		SELECT																				"
		f_SQL_____ = f_SQL_____ & "			ROW_NUMBER() OVER (ORDER BY C.NO_ORDER ASC, A.NM_STORE ASC) AS [No]								"
		f_SQL_____ = f_SQL_____ & "			, A.CD_STOREID, B.NM_CODE AS CD_BRANDID, A.CD_STORECD, C.NM_CODE AS NM_STORETYPE, A.NO_VIEWORDER, D.NM_CODE AS NM_AREA, A.NM_STORE, A.NM_ADDR, A.NO_TEL1, C.NO_ORDER "
		f_SQL_____ = f_SQL_____ & "		FROM																				"
		f_SQL_____ = f_SQL_____ & "			T_STORE AS A																	"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS B ON B.CD_GUBUN = 'BRANDID' AND A.CD_BRANDID = B.CD_CODE		"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS C ON C.CD_GUBUN = 'STORETYPE' AND A.NM_STORETYPE = C.CD_CODE	"
		f_SQL_____ = f_SQL_____ & "				INNER JOIN T_CODE AS D ON D.CD_GUBUN = 'AREA' AND A.NM_AREA = D.CD_CODE				"	
		f_SQL_____ = f_SQL_____ & "		WHERE A.YN_USE = 'Y' AND C.YN_USE = 'Y' AND A.YN_ING = 'Y' AND B.CD_CODE = '" & p_cd_brandid___ & "'"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery
		f_SQL_____ = f_SQL_____ & "	) AS J									"
		f_SQL_____ = f_SQL_____ & "WHERE										"
		f_SQL_____ = f_SQL_____ & "	No BETWEEN @SNUM AND @ENUM				"
		f_SQL_____ = f_SQL_____ & "ORDER BY									"
		f_SQL_____ = f_SQL_____ & "	NO_ORDER ASC				"

		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_StoreList_paging = f_AryList_____
	End Function


	Function fn_StoreList_paging_Union (p_cd_brandid___, p_page___, p_pgsize___, p_AddWhereQuery)
		Dim f_objADO_____, f_Rs_____, f_SQL_____
		Dim f_AryList_____

		Set f_objADO_____ = new clsADO
		

		f_SQL_____ = f_SQL_____ & " DECLARE @CNT int									"								
		f_SQL_____ = f_SQL_____ & " DECLARE @SNUM int									"			
		f_SQL_____ = f_SQL_____ & " DECLARE @ENUM int									"			
		f_SQL_____ = f_SQL_____ & "SET @ENUM = " & p_page___ & " * " & p_pgsize___ & "	"
		f_SQL_____ = f_SQL_____ & "SET @SNUM = (@ENUM - " & p_pgsize___ & ")+1			"		
													
		f_SQL_____ = f_SQL_____ & " SELECT @CNT = COUNT(*)								"
		f_SQL_____ = f_SQL_____ & " FROM												"
		f_SQL_____ = f_SQL_____ & " 		T_STORE AS A																												"
		f_SQL_____ = f_SQL_____ & " 			INNER JOIN T_CODE AS B ON B.CD_GUBUN = 'BRANDID' AND B.CD_CODE = A.CD_BRANDID 											"
		f_SQL_____ = f_SQL_____ & " 			INNER JOIN T_CODE AS C ON C.CD_GUBUN = 'STORETYPE' AND C.CD_CODE = A.NM_STORETYPE										"
		f_SQL_____ = f_SQL_____ & " 			INNER JOIN T_CODE AS D ON D.CD_GUBUN = 'AREA' AND D.CD_CODE = A.NM_AREA													"
		f_SQL_____ = f_SQL_____ & " 	WHERE A.YN_USE = 'Y' AND A.YN_ING = 'Y'																							"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery						
		
		f_SQL_____ = f_SQL_____ & " SELECT												"														
		f_SQL_____ = f_SQL_____ & " 	@CNT AS CNT, NO, CD_STOREID, CD_BRANDID, CD_STORECD, NM_STORETYPE, NO_VIEWORDER, NM_AREA, NM_STORE, NM_ADDR, NO_TEL1, NO_ORDER  "
											'0		'1		'2			'3			'4			'5				'6			'7		'8			'9		'10		'11
		f_SQL_____ = f_SQL_____ & " FROM																															    "
		f_SQL_____ = f_SQL_____ & " (																																	"
		f_SQL_____ = f_SQL_____ & " 	SELECT																															"
		f_SQL_____ = f_SQL_____ & " 		ROW_NUMBER() OVER (ORDER BY B.NO_ORDER ASC, C.NO_ORDER ASC, B.NM_CODE ASC, A.NM_STORE ASC) AS [No]											"
		f_SQL_____ = f_SQL_____ & " 		, A.CD_STOREID, B.NM_CODE AS CD_BRANDID, A.CD_STORECD, C.NM_CODE AS NM_STORETYPE, A.NO_VIEWORDER, D.NM_CODE AS NM_AREA, A.NM_STORE, A.NM_ADDR, A.NO_TEL1, C.NO_ORDER "
		f_SQL_____ = f_SQL_____ & " 	FROM																															"
		f_SQL_____ = f_SQL_____ & " 		T_STORE AS A																												"
		f_SQL_____ = f_SQL_____ & " 			INNER JOIN T_CODE AS B ON B.CD_GUBUN = 'BRANDID' AND B.CD_CODE = A.CD_BRANDID 											"
		f_SQL_____ = f_SQL_____ & " 			INNER JOIN T_CODE AS C ON C.CD_GUBUN = 'STORETYPE' AND C.CD_CODE = A.NM_STORETYPE										"
		f_SQL_____ = f_SQL_____ & " 			INNER JOIN T_CODE AS D ON D.CD_GUBUN = 'AREA' AND D.CD_CODE = A.NM_AREA													"
		f_SQL_____ = f_SQL_____ & " 	WHERE A.YN_USE = 'Y' AND A.YN_ING = 'Y'																							"
		f_SQL_____ = f_SQL_____ & p_AddWhereQuery		
		f_SQL_____ = f_SQL_____ & " ) AS J																																"					
		f_SQL_____ = f_SQL_____ & " WHERE																																"
		f_SQL_____ = f_SQL_____ & " 	No BETWEEN @SNUM AND @ENUM																										"
		f_SQL_____ = f_SQL_____ & " ORDER BY																															"
		f_SQL_____ = f_SQL_____ & " 	NO_ORDER ASC, CD_BRANDID ASC,  NM_STORE ASC																						"
			
		f_objADO_____.setSql(f_SQL_____)
		f_AryList_____ = f_objADO_____.getArrRs()

		Set f_objADO_____ = Nothing

		fn_StoreList_paging_Union = f_AryList_____
	End Function
%>