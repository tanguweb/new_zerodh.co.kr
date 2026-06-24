<!-- #Include virtual = "/Include/config.asp" -->
<!-- #include virtual="/function/fn_AdminLoginCheck.asp" -->
<%
	Dim CD_BOARDCD : CD_BOARDCD = ReqQ("CD_BOARDCD")
	Dim sYN_PERIOD : sYN_PERIOD = ReqQ("sYN_PERIOD")	'기간검색 사용여부
	Dim DT_SYMD : DT_SYMD = ReqQ("DT_SYMD")	'기간검색(시작일)
	Dim DT_EYMD : DT_EYMD = ReqQ("DT_EYMD")	'기간검색(종료일)
	
	Dim sYN_DISPLAY : sYN_DISPLAY = ReqQ("sYN_DISPLAY")	'노출여부
	Dim sGUBUN : sGUBUN = ReqQ("sGUBUN")
	Dim sKEYWORD : sKEYWORD = ReqQ("sKEYWORD")
	Dim sSTATUS : sSTATUS = ReqQ("sSTATUS") '답변 처리 상태
'	Dim sCOUNSEL_TYPE : sCOUNSEL_TYPE = ReqQ("sCOUNSEL_TYPE")	'문의 분류
	Dim sPUBLIC : sPUBLIC = ReqQ("sPUBLIC") '상담글 공개 여부(공개 : Y / 비공개 : N)

	Dim objADO, SQL, iloop, jloop, kloop 
	Dim aryBoardInfo, AryRs
	Dim NM_BOARDNM, YN_USE, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, YN_DOWNLOAD, YN_TAG
	Dim spancnt : spancnt = 6

	If CD_BOARDCD = "" Then
		Response.Write "<script>alert('올바른경로로 접근하세요.');history.back();</script>"
		Response.End
	End If
	
	' init(검색 시작일)
	If DT_SYMD = "" Then 
		DT_SYMD = Left(Now(),10)
	End If 
	
	' init(검색 종료일)
	If DT_EYMD = "" Then 
		DT_EYMD = Left(Now(),10)
	End If

	
	
	' init(검색 전시여부)
	If sYN_DISPLAY = "" Then 
		sYN_DISPLAY = "Y"
	End If

	

	Set objADO = new clsADO
		
		

		SQL= ""
		
		SQL = SQL & "		SELECT																								"
		SQL = SQL & "		ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]	"
		SQL = SQL & "		, BM.NM_BOARDNM, BM.CD_BOARDCD, BM.YN_USE, BM.YN_CNTVIEW, BM.YN_ANSWER, BM.YN_REPLY, BM.YN_XML		"

		SQL = SQL & "		, BD.CD_BOARDID, BD.NO_BOARD_DEPTH, BD.NO_BOARD_SEQ, BD.YN_DISPLAY, BD.NO_VIEWCNT, BD.NO_VIEWORDER	"

		SQL = SQL & "		, BD.NM_TITLE, BD.NM_CONTENTS, BD.CD_INUSER, BD.DT_INSYSDATE										"

		SQL = SQL & "		, (SELECT COUNT(*) FROM T_BOARDREPLY WHERE CD_BOARDID = BD.CD_BOARDID AND YN_USE = 'Y') AS NO_REPLY	"	
		SQL = SQL & "		, BD.NO_DOWNLOAD, BD.CD_BOARDKEY, AM.YN_LOGIN, AM.NM_NAME, BD.NM_FIELD_15							"

		SQL = SQL & "		, ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER DESC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ DESC) AS [NoDesc]"	
		SQL = SQL & "		, BD.YN_USE AS [BD.YN_USE]				"	
		SQL = SQL & "		, CD2.CD_CODE AS [CD2.CD_CODE], CD2.NM_CODE AS [CD2.NM_CODE]										"	
		SQL = SQL & "		, BD.NM_FIELD_6																						"
		SQL = SQL & "		FROM																								"
		SQL = SQL & "			T_BOARDMAST AS BM																				"
		SQL = SQL & "				INNER JOIN T_BOARDDETAIL AS BD																"
		SQL = SQL & "				ON BD.CD_BOARDCD = BM.CD_BOARDCD															"
		SQL = SQL & "				INNER JOIN T_ADMIN AS AM																	"
		SQL = SQL & "				ON AM.CD_ADMINID = BD.CD_INUSER																"
		
		SQL = SQL & "				INNER JOIN T_CODE AS CD2																	"
		SQL = SQL & "				ON CD2.CD_CODE = BD.NM_FIELD_11																"
		SQL = SQL & "		WHERE                                                                                               "
		SQL = SQL & "				BM.CD_BOARDCD = '" & CD_BOARDCD & "' AND CD2.CD_GUBUN = 'STATUS'"
		SQL = SQL & "			AND BD.YN_USE = 'Y'														"

' 키워드 검색 조건
If sKEYWORD <> "" Then
	If sGUBUN = "title" Then
		SQL = SQL & "			AND BD.NM_TITLE LIKE '%'+'" & sKEYWORD & "'+'%'											"
	ElseIf sGUBUN = "contents" Then 
		SQL = SQL & "			AND BD.NM_CONTENTS LIKE '%'+'" & sKEYWORD & "'+'%'										"
	ElseIf sGUBUN = "userid" Then 
		SQL = SQL & "			AND BD.CD_INUSER = '" & sKEYWORD & "'													"
	ElseIf sGUBUN = "usernm" Then 
		SQL = SQL & "			AND AM.NM_NAME = '" & sKEYWORD & "'														"
	End If 
End If
	
' 기간 검색 조건
If sYN_PERIOD = "Y" Then 
		SQL = SQL & "			AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) >= " & Replace(DT_SYMD,"-","") & "		"
		SQL = SQL & "			AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) <= " & Replace(DT_EYMD,"-","") & "		"
End If 



' 전시여부 조건
If sYN_DISPLAY <> "" And sYN_DISPLAY <> "ALL" Then 
		SQL = SQL & "			AND BD.YN_DISPLAY = '" & sYN_DISPLAY & "'												"
End If 

' 답변 처리상태 조건
If sSTATUS <> "" And sSTATUS <> "ALL" Then 
		SQL = SQL & "			AND BD.NM_FIELD_11 = '" & sSTATUS & "'													"
End If

' 공개여부
If sPUBLIC <> "" Then 
		SQL = SQL & "			AND BD.NM_FIELD_3 = '" & sPUBLIC & "'													"
End If


		SQL = SQL & "ORDER BY																									"
		SQL = SQL & "	NO ASC																									"

	objADO.setSql(SQL)
	AryRs = objADO.getArrRs()

	Set objADO = Nothing 

	Response.buffer = True 
	Response.CharSet ="utf-8"
	Response.contenttype = "application/vnd.ms-excel" 
	Response.AddHeader "Content-Disposition","attachment;filename=" & Server.UrlEncode("온라인상담") & "_" & Replace(Date,"-","") & ".xls"
%>
<HTML>
 <BODY>
	<table cellSpacing="1" cellPadding="1" width="100%" bgcolor="#000" align="left" border="1">
	<tr align="center" bgColor="#ffffff" height="30">
		<td align="center" width="*"><b>No</b></td>
		
		<td align="center" width="*"><b>제목</b></td>
		<td align="center" width="*"><b>내용</b></td>
		<td align="center" width="*"><b>작성자</b></td>
		<td align="center" width="*"><b>처리상태</b></td>
		<td align="center" width="*"><b>조회수</b></td>
		<td align="center" width="*"><b>사용여부</b></td>
		<td align="center" width="*"><b>전시여부</b></td>
		
		<td align="center" width="*"><b>등록일</b></td>
	</tr>
	<%
	If IsArray(AryRs) Then
		For iloop = 0 To UBound(AryRs,2)
	%>
	<tr height="25" bgcolor="#ffffff">
		<td align="center" width="*"><%=AryRs(24,iloop)%></td>
		
		<td align="center" width="*"><%=AryRs(14,iloop)%></td>
		<td align="center" width="*"><%=AryRs(15,iloop)%></td>
		<td align="center" width="*"><%=AryRs(28,iloop)%></td>
		<td align="center" width="*"><%=AryRs(27,iloop)%></td>
		<td align="center" width="*"><%=AryRs(12,iloop)%></td>
		<td align="center" width="*"><%=AryRs(25,iloop)%></td>
		<td align="center" width="*"><%=AryRs(11,iloop)%></td>
		<td align="center" width="*"><%=AryRs(17,iloop)%></td>
	</tr>
	<%
		Next 
	End If 
	%>
	</table>

</BODY>
</HTML>