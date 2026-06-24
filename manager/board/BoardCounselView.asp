<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<!-- #include virtual = "/function/fn_FileManager.asp" -->
<!-- #include virtual="/admin/include/header.asp" -->
<!-- #Include Virtual = "/FCKeditor/fckeditor.asp" -->
<%
	'########################################################################################
	'#	File		: /admin/board/BoardCounselView.asp
	'#  Create		: / 2011.12.16
	'#	Info		: 1:1문의 답변 작성 페이지
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	
	Dim CD_BOARDID : CD_BOARDID = ReqQ("CD_BOARDID")

	Dim CD_BOARDCD, CD_BOARDKEY, NO_BOARD_DEPTH, NO_BOARD_SEQ
	Dim aryCodeInfo, aryBoardInfo, iloop, jloop, kloop
	Dim objADO, SQL
	Dim NM_BOARDNM, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, YN_DOWNLOAD
	Dim YN_DISPLAY,  NO_VIEWCNT, NO_VIEWORDER, NO_DOWNLOAD, YN_USE, NM_TITLE, NM_CONTENTS, CD_INUSER, DT_INSYSDATE
	Dim NM_XMLPATH, NM_XMLPATH2, NM_XMLQUERY, NM_XMLFILE, YN_TAG, NM_TAG
	Dim ary_FIELD(16), ary_FIELDNM(16), ary_FIELDINFO(16), ary_FIELDTYPE(16), ary_FIELDOPTION(16)
	Dim ary_STATUS, ary_COUNSEL
	Dim NM_FIELD_1, NM_FIELD_2, NM_FIELD_3, NM_FIELD_4, NM_FIELD_5, NM_FIELD_6, NM_FIELD_7
	Dim NM_FIELD_8, NM_FIELD_9, NM_FIELD_10, NM_FIELD_11, NM_FIELD_12, NM_FIELD_13, NM_FIELD_14, NM_FIELD_15
	Dim CD_BOARDID_reply, NM_TITLE_reply, NM_CONTENTS_reply, NM_NAME_reply, DT_INSYSDATE_reply, ary_Answer, MODE

	Set objADO = new clsADO
	
	SQL = ""
	SQL = SQL & "SELECT																									"
	SQL = SQL & "	BM.NM_BOARDNM, BM.YN_USE, BM.YN_CNTVIEW, BM.YN_ANSWER, BM.YN_REPLY, BM.YN_XML						"	'5
	SQL = SQL & "	, BM.NM_FIELDNM_1, BM.NM_FIELDNM_2, BM.NM_FIELDNM_3, BM.NM_FIELDNM_4, BM.NM_FIELDNM_5				"	'10
	SQL = SQL & "	, BM.NM_FIELDNM_6, BM.NM_FIELDNM_7, BM.NM_FIELDNM_8, BM.NM_FIELDNM_9, BM.NM_FIELDNM_10				"	'15
	SQL = SQL & "	, BM.NM_FIELDNM_11, BM.NM_FIELDNM_12, BM.NM_FIELDNM_13, BM.NM_FIELDNM_14, BM.NM_FIELDNM_15			"	'20

	SQL = SQL & "	, BM.NM_FIELDINFO_1, BM.NM_FIELDINFO_2, BM.NM_FIELDINFO_3, BM.NM_FIELDINFO_4, BM.NM_FIELDINFO_5		"	'25
	SQL = SQL & "	, BM.NM_FIELDINFO_6, BM.NM_FIELDINFO_7, BM.NM_FIELDINFO_8, BM.NM_FIELDINFO_9, BM.NM_FIELDINFO_10	"	'30
	SQL = SQL & "	, BM.NM_FIELDINFO_11, BM.NM_FIELDINFO_12, BM.NM_FIELDINFO_13, BM.NM_FIELDINFO_14, BM.NM_FIELDINFO_15"	'35

	SQL = SQL & "	, BM.CD_FIELDTYPE_1, BM.CD_FIELDTYPE_2, BM.CD_FIELDTYPE_3, BM.CD_FIELDTYPE_4, BM.CD_FIELDTYPE_5		"	'40
	SQL = SQL & "	, BM.CD_FIELDTYPE_6, BM.CD_FIELDTYPE_7, BM.CD_FIELDTYPE_8, BM.CD_FIELDTYPE_9, BM.CD_FIELDTYPE_10	"	'45
	SQL = SQL & "	, BM.CD_FIELDTYPE_11, BM.CD_FIELDTYPE_12, BM.CD_FIELDTYPE_13, BM.CD_FIELDTYPE_14, BM.CD_FIELDTYPE_15"	'50

	SQL = SQL & "	, BD.CD_BOARDID, BD.NO_BOARD_DEPTH, BD.NO_BOARD_SEQ, BD.CD_BOARDCD, BD.YN_DISPLAY					"	'55
	SQL = SQL & "	, BD.NO_VIEWCNT, BD.NO_VIEWORDER, BD.YN_USE, BD.NM_TITLE, BD.NM_CONTENTS							"	'60

	SQL = SQL & "	, BD.NM_FIELD_1, BD.NM_FIELD_2, BD.NM_FIELD_3, BD.NM_FIELD_4, BD.NM_FIELD_5							"	'65
	SQL = SQL & "	, BD.NM_FIELD_6, BD.NM_FIELD_7, BD.NM_FIELD_8, BD.NM_FIELD_9, BD.NM_FIELD_10						"	'70
	SQL = SQL & "	, BD.NM_FIELD_11, BD.NM_FIELD_12, BD.NM_FIELD_13, BD.NM_FIELD_14, BD.NM_FIELD_15					"	'75

	SQL = SQL & "	, BD.CD_INUSER, BD.DT_INSYSDATE, BD.CD_BOARDKEY, BM.YN_DOWNLOAD, BD.NO_DOWNLOAD						"	'80
	SQL = SQL & "	, BM.NM_XMLPATH, BM.NM_XMLQUERY, BM.NM_XMLFILE, BM.YN_TAG, BD.NM_TAG								"	'85

	SQL = SQL & "	, BM.YN_FIELDOPTION_1, BM.YN_FIELDOPTION_2, BM.YN_FIELDOPTION_3, BM.YN_FIELDOPTION_4				"	'89
	SQL = SQL & "	, BM.YN_FIELDOPTION_5, BM.YN_FIELDOPTION_6, BM.YN_FIELDOPTION_7, BM.YN_FIELDOPTION_8				"	'93
	SQL = SQL & "	, BM.YN_FIELDOPTION_9, BM.YN_FIELDOPTION_10, BM.YN_FIELDOPTION_11, BM.YN_FIELDOPTION_12				"	'97
	SQL = SQL & "	, BM.YN_FIELDOPTION_13, BM.YN_FIELDOPTION_14, BM.YN_FIELDOPTION_15									"	'100
	SQL = SQL & "FROM																									"
	SQL = SQL & "	T_BOARDDETAIL AS BD																					"
	SQL = SQL & "		INNER JOIN T_BOARDMAST AS BM ON BM.CD_BOARDCD = BD.CD_BOARDCD									"
	SQL = SQL & "WHERE																									"
	SQL = SQL & "	BD.YN_USE = 'Y' AND BD.CD_BOARDID = '" & CD_BOARDID & "' 											"
	objADO.setSql(SQL)
	aryBoardInfo = objADO.getArrRs()

	If not IsArray(aryBoardInfo) Then
		Response.Write "<script>alert('올바른 경로로 접근하세요.');history.back();</script>"
		Response.End
	End If
	
	CD_BOARDCD		= aryBoardInfo(54,0)
	CD_BOARDKEY		= aryBoardInfo(78,0)
	NO_BOARD_DEPTH	= aryBoardInfo(52,0)
	NO_BOARD_SEQ	= aryBoardInfo(53,0)

	NM_BOARDNM		= aryBoardInfo(0,0)
	YN_CNTVIEW		= aryBoardInfo(2,0)
	YN_ANSWER		= aryBoardInfo(3,0)
	YN_REPLY		= aryBoardInfo(4,0)
	YN_XML			= aryBoardInfo(5,0)
	YN_DOWNLOAD		= aryBoardInfo(79,0)

	YN_DISPLAY		= aryBoardInfo(55,0)
	NO_VIEWCNT		= aryBoardInfo(56,0)
	NO_VIEWORDER	= aryBoardInfo(57,0)
	NO_DOWNLOAD		= aryBoardInfo(80,0)
	YN_USE			= aryBoardInfo(58,0)
	NM_TITLE		= aryBoardInfo(59,0)
	NM_CONTENTS		= aryBoardInfo(60,0)
	CD_INUSER		= aryBoardInfo(76,0)
	DT_INSYSDATE	= aryBoardInfo(77,0)

	NM_XMLPATH		= aryBoardInfo(81,0)
	NM_XMLPATH2		= "/" & replace(aryBoardInfo(81,0),"\","/")
	NM_XMLQUERY		= aryBoardInfo(82,0)

	YN_TAG			= aryBoardInfo(84,0)
	NM_TAG			= aryBoardInfo(85,0)
	
	NM_FIELD_1		= aryBoardInfo(61,0)
	NM_FIELD_2		= aryBoardInfo(62,0)
	NM_FIELD_3		= aryBoardInfo(63,0)
	NM_FIELD_4		= aryBoardInfo(64,0)
	NM_FIELD_5		= aryBoardInfo(65,0)
	NM_FIELD_6		= aryBoardInfo(66,0)
	NM_FIELD_7		= aryBoardInfo(67,0)
	NM_FIELD_8		= aryBoardInfo(68,0)
	NM_FIELD_9		= aryBoardInfo(69,0)
	NM_FIELD_10		= aryBoardInfo(70,0)
	NM_FIELD_11		= aryBoardInfo(71,0)
	NM_FIELD_12		= aryBoardInfo(72,0)
	NM_FIELD_13		= aryBoardInfo(73,0)
	NM_FIELD_14		= aryBoardInfo(74,0)
	NM_FIELD_15		= aryBoardInfo(75,0)

	'// 문의구분 가져오기
'	ary_COUNSEL = fn_CodeInfoWhere("CLINICTYPE"," AND CD_CODE = '" & NM_FIELD_1 & "'")

	'// 처리상태 가져오기
	ary_STATUS = fn_CodeInfoWhere("STATUS"," AND CD_CODE = '" & NM_FIELD_11 & "'")

	'// 필드 도움말 가져오기
	For iloop = 1 To 15
		ary_FIELDINFO(iloop)	= aryBoardInfo(iloop+20,0)
	Next 
	
	'// 답변글 가져오기(무조건 1개)
	SQL = ""
	SQL = SQL & "	SELECT TOP 1																					"
	SQL = SQL & "		BD.CD_BOARDID, BD.CD_BOARDKEY, BD.NO_BOARD_DEPTH, BD.NO_BOARD_SEQ, BD.CD_BOARDCD, BD.YN_USE	"	'5
	SQL = SQL & "		, BD.YN_ADMIN, BD.NM_TITLE, BD.NM_CONTENTS, BD.CD_INUSER, BD.DT_INSYSDATE					"	'10
	SQL = SQL & "		, AD.NM_NAME																				"	'11
	SQL = SQL & "	FROM T_BOARDDETAIL AS BD																		"
	SQL = SQL & "		INNER JOIN T_ADMIN AS AD																	"
	SQL = SQL & "		ON AD.CD_ADMINID = BD.CD_INUSER																"
	SQL = SQL & "	WHERE CD_BOARDKEY = '" & CD_BOARDID & "' AND NO_BOARD_DEPTH = 2 AND NO_BOARD_SEQ = 2		"

	objADO.setSql(SQL)
	ary_Answer = objADO.getArrRs()
	
	If IsArray(ary_Answer) Then 
		MODE = "U"
		CD_BOARDID_reply = ary_Answer(0,0)
		NM_CONTENTS_reply = ary_Answer(8,0)
		NM_NAME_reply = ary_Answer(11,0)
		DT_INSYSDATE_reply = ary_Answer(10,0)
	Else 
		MODE = "I"
		'NM_CONTENTS_reply = "안녕하세요. "&NM_FIELD_6&"님^^<br />"
		'NM_CONTENTS_reply = NM_CONTENTS_reply & "여성의 건강과 행복을 가장 먼저 생각하는 여성전문병원 “이혜미 산부인과”입니다.<br /><br />"
		'NM_CONTENTS_reply = NM_CONTENTS_reply & "방문하시면 보다 더 자세하고 정확한 상담을 받으실 수 있습니다.<br />"
		'NM_CONTENTS_reply = NM_CONTENTS_reply & "예약 전화는 Tel) 02-508-4693 입니다.<br /><br />"
		'NM_CONTENTS_reply = NM_CONTENTS_reply & "감사합니다. 언제나 행복한 하루 보내세요<br /><br />"
		'NM_CONTENTS_reply = NM_CONTENTS_reply & "오시는길: 선릉역 10번 출구 앞 리치타워 301호"
	End If 	

	Set objADO = Nothing

%>
	<script type="text/Javascript">
		/**
		 *
		 */
		function fn_answer()
		{
			var oEditor = FCKeditorAPI.GetInstance("NM_CONTENTS_reply");

			if (oEditor.GetXHTML(true) == "") {

				alert("답변내용을 입력하세요.");
				oEditor.Focus();
				return;
			}
			
			document.frm_boardins.submit();
		}
		
		

		function fn_del()
		{
			if(confirm("질문내용과 답변이 모두 삭제됩니다.\n\n삭제 하시겠습니까?") == true)
			{
				document.getElementById("MODE").value = "D";
				document.getElementById("YN_USE").value = "N";
				fn_answer();
			}
			else
			{
				return;
			}
		}

		function fn_toggleHelper()
		{
			$("#adminHelper").toggle();
		}

		/**
		 *	각 필드 도움말 표시(모두 숨겼다가, 해당 필드 도움말만 표시)
		 */
		$(document).ready(function(){
			$(".field_info").hover(
				function(){
					$(".field_info_txt").hide();
					$(this).next(".field_info_txt").show();
				},
				function(){
					$(this).next(".field_info_txt").hide();
			});
		});


		/**
		 *	인쇄
		 */
		function fn_Print()
		{
			if (confirm("세로 인쇄시 우측이 잘릴수 있습니다.\n\n인쇄를 하시겠습니까?") == true)
			{
				window.print();
			}
		}
	</Script>
</head>
<body>
<form id="frm_boardins" name="frm_boardins" method="post" action="BoardCounsel_Proc.asp" ENCTYPE="multipart/form-data">
	<input type="hidden" id="MODE" name="MODE" value="<%=MODE%>">
	<input type="hidden" id="CD_BOARDCD" name="CD_BOARDCD" value="<%=CD_BOARDCD%>">
	<input type="hidden" id="CD_BOARDID" name="CD_BOARDID" value="<%=CD_BOARDID%>">
	<input type="hidden" id="CD_BOARDID_reply" name="CD_BOARDID_reply" value="<%=CD_BOARDID_reply%>">
	<input type="hidden" id="CD_BOARDKEY" name="CD_BOARDKEY" value="<%=CD_BOARDKEY%>">
	<input type="hidden" id="NO_BOARD_DEPTH" name="NO_BOARD_DEPTH" value="2">
	<input type="hidden" id="NO_BOARD_SEQ" name="NO_BOARD_SEQ" value="2">
	<input type="hidden" id="YN_XML" name="YN_XML" value="<%=YN_XML%>">
	<input type="hidden" id="NM_XMLPATH" name="NM_XMLPATH" value="<%=NM_XMLPATH%>">
	<input type="hidden" id="NM_XMLFILE" name="NM_XMLFILE" value="<%=NM_XMLFILE%>">
	<input type="hidden" id="YN_USE" name="YN_USE" value="Y">
	<input type="hidden" id="NO_DOWNLOAD" name="NO_DOWNLOAD" value="<%=NO_DOWNLOAD%>">
	<input type="hidden" id="NM_TAG" name="NM_TAG" value="<%=NM_TAG%>">
	
	<input type="hidden" id="emailTo" name="emailTo" value="<%=NM_FIELD_7%>"> <!-- 답변시 받는 사람 메일(고객) 주소 -->

	<div id="wrap">
		<div class="section">
			<board_header>
			<div class="board_header"> 
				<div class="board_title"> 
					<%=NM_BOARDNM%> 수정
				</div> 
			</div> 
			</board_header>
	
			<board_section>
			<div class="board_section">
				<div class="board_table_section"> 	
					<table class="board_table">
						<tbody>
							<tr>
								<!--
								<th>상담분류</th>
								<td>
									<%'=ary_COUNSEL(2,0)%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%'=ary_FIELDINFO(1)%></div>
								</td>
								-->
								<th>처리상태</th>
								<td>
									<%=ary_STATUS(2,0)%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(11)%></div>
								</td>
								<th>등록일</th>
								<td><%=DT_INSYSDATE%></td>
								<th></th>
								<td></td>
							</tr>
							<tr>
								<th>고객명</th>
								<td>
									<%=NM_FIELD_6%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(6)%></div>
								</td>
								<th>이메일</th>
								<td>
									<%=NM_FIELD_7%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(7)%></div>
								</td>
								<th>연락처</th>
								<td>
									<%=NM_FIELD_8%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(8)%></div>
								</td>
							</tr>
					<%if CD_BOARDCD="1005" then%>
							<tr>
						
								<th>비밀번호</th>
								<td>
									<%=NM_FIELD_4%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(4)%></div>
								</td>
							
								<th>첨부파일</th>
								<td colspan="3">
									<%If NM_FIELD_5 <> "" Then%>
									<a href="/file/Board/<%=NM_FIELD_5%>" target="_blank"><%=GetFileExtImage(NM_FIELD_5)%><%=NM_FIELD_5%></a>
									<%Else%>
									첨부파일이 없습니다.
									<%End If%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(5)%></div>
								</td>
							</tr>
					
						
							<tr>
								<th>답변방법</th>
								<td>
									<%=NM_FIELD_2%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(2)%></div>
								</td>
								<th>공개여부</th>
								<td colspan="3">
									<%=getBoolean(NM_FIELD_3 = "Y","공개","비공개")%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(3)%></div>
								</td>
							</tr>
					<%end if%>
						<!--	<tr>
								<th>평균생리주기</th>
								<td>
									<%=NM_FIELD_9%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(9)%></div>
								</td>
								<th>마지막생리일</th>
								<td colspan="3">
									<%=NM_FIELD_10%>
									<span class="field_info"></span>
									<div class="field_info_txt"><%=ary_FIELDINFO(10)%></div>
								</td>
							</tr>-->

							<input type="hidden" id="YN_DISPLAY" name="YN_DISPLAY" value="<%=YN_DISPLAY%>" />
							<input type="hidden" id="NO_VIEWORDER" name="NO_VIEWORDER" value="<%=NO_VIEWORDER%>" />
							<input type="hidden" id="NO_CNTVIEW" name="NO_CNTVIEW" value="<%=NO_VIEWCNT%>">
							<!--
							<tr>
								<th>전시여부</th>
								<td>
									<select class="select" id="YN_DISPLAY" name="YN_DISPLAY">
										<option value="Y" <%=getBoolean(YN_DISPLAY="Y","selected","")%>>예</option>
										<option value="N" <%=getBoolean(YN_DISPLAY="N","selected","")%>>아니오</option>
									</select>
									<span class="field_info"></span>
									<div class="field_info_txt">프론트 페이지에 전시여부를 선택합니다.</div>
								</td>
								<th>TOP 공지</th>
								<td>
									<%if NO_BOARD_DEPTH = 1 then %>
										<select class="select" id="NO_VIEWORDER" name="NO_VIEWORDER">
											<option value="100" <%=getBoolean(NO_VIEWORDER="100","selected","")%>>아니오</option>
											<option value="0" <%=getBoolean(NO_VIEWORDER="0","selected","")%>>예</option>
										</select>
										<span class="field_info"></span>
										<div class="field_info_txt">공지글은 언제나 최상단에 위치합니다.</div>
									<%else %>
										<input type="hidden" id="NO_VIEWORDER" name="NO_VIEWORDER" value="<%=NO_VIEWORDER%>">
									<%end if %>									
								</td>
								<th>조회수</th>
								<td>
									<%if YN_CNTVIEW = "Y" then %>
										<input class="input" type="text" id="NO_CNTVIEW" name="NO_CNTVIEW" value="<%=NO_VIEWCNT%>" maxlength="5" size="5">
										<span class="field_info"></span>
										<div class="field_info_txt">조회수를 변경할 수 있습니다.</div>
									<%else %>
										<input type="hidden" id="NO_CNTVIEW" name="NO_CNTVIEW" value="<%=NO_VIEWCNT%>">
									<%end if %>
								</td>
							</tr>
							-->
							<tr>
								<th>문의제목</th>
								<td colspan="5"><%=NM_TITLE%></td>
							</tr>
							<tr>
								<th>문의내용</th>
								<td colspan="5">
									<div class="box_counsel">
										<%=NM_CONTENTS%>
									</div>	
								</td>
							</tr>
							<tr>
								<th>답변내용</th>
								<td colspan="5"><%Editor "NM_CONTENTS_reply", "100%", "400", NM_CONTENTS_reply%></td>
							</tr>
							<tr>
								<th>이메일전송여부</th>
								<td colspan="5">
									&nbsp;
									<input type="radio" class="radio" id="YN_EMAIL" name="YN_EMAIL" value="Y" />예&nbsp;
									<input type="radio" class="radio" id="YN_EMAIL" name="YN_EMAIL" value="N" checked/>아니오
								</td>
							</tr>
							<tr>
								<th>처리자</th>
								<td><%=NM_NAME_reply%></td>
								<th>처리일</th>
								<td colspan="3"><%=DT_INSYSDATE_reply%></td>
							</tr>
						</tbody>
						<tfoot>
							<tr>
								<td colspan="6">
									<a href="javascript:fn_answer();" class="btn_image" title="답변하기"><span class="reply">답변하기</span></a>
									<!-- <a href="javascript:fn_save();" class="btn_image" title="저장"><span class="save">저장</span></a> -->
									<a href="javascript:fn_del();" class="btn_image" title="삭제"><span class="delete">삭제</span></a>
									<!--<a href="javascript:fn_Print();" class="btn_image" title="인쇄"><span class="print">인쇄</span></a>-->
									<a href="javascript:document.location.href='BoardCounsel.asp?CD_BOARDCD=<%=CD_BOARDCD%>';" class="btn_image" title="목록으로 돌아가기"><span class="list">목록으로 돌아가기</span></a>
									<a href="javascript:document.location.reload();" class="btn_image" title="새로고침"><span class="reload">새로고침</span></a>
									<a href="javascript:fn_toggleHelper();" class="btn_image" title="도움말 보기"><span class="help">도움말 보기</span></a>
								</td>
							</tr>
							
							<!-- str: 도움말 시작 -->
							<!-- #include virtual = "/admin/Board/AdminHelper.asp" -->
							<!-- end: 도움말 끝 -->

						</tfoot>
					</table>

				</div> 
			</div> 
			</board_section>

		</div>
	</div>


</form>
</body>
</html>