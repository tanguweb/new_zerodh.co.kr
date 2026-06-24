<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_PageNavigation.asp"-->
<%
	'########################################################################################
	'#	File		: /manager/board/BoardDetail.asp
	'#  Create		: / 2010.09.08
	'#	Info		: 게시판글리스트
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim aryBoardList, aryBoardInfo
	Dim objADO, SQL, iloop, jloop, kloop 
	Dim CD_BOARDCD, sGUBUN, sKEYWORD, sDISPLAY, sSTARTDATE, sENDDATE
	Dim NM_BOARDNM, YN_USE, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, YN_DOWNLOAD, CD_BOARDTYPE
	Dim spancnt : spancnt = 5

	CD_BOARDCD = ReqQ("CD_BOARDCD")
	sGUBUN = Req("sGUBUN")
	sKEYWORD = Req("sKEYWORD")
	sSTARTDATE = Req("sSTARTDATE")
	sENDDATE = Req("sENDDATE")
	sDISPLAY = Req("sDISPLAY")

	if CD_BOARDCD = "" then 
		response.write "<script>alert('올바른경로로 접근하세요.');history.back();</script>"
		response.end 
	end if 
	
	'If sSTARTDATE = "" Then
	'	sSTARTDATE = Date()	'검색 시작일
	'End If

	'If sENDDATE = "" Then
	'	sENDDATE = Date()	'검색 종료일
	'End If

	Dim page, startpage, tn, pgcount, nextpage, prevpage
	Const pgsize = 10	
	Const setsize = 10

	If IsNumeric(ReqQ("page")) = False then  
		page=1  
		startpage=1
	Else 
		page=cint(ReqQ("page"))
		startpage=int(page/setsize)
			If startpage=(page/setsize) Then 
				startpage=page-setsize + 1
			Else 
				startpage=int(page/setsize)*setsize + 1
			End If 
	End If 

	Set objADO = new clsADO
		
		SQL = ""
		SQL = SQL & "SELECT BM.NM_BOARDNM, BM.YN_USE, BM.YN_CNTVIEW, BM.YN_ANSWER, BM.YN_REPLY, BM.YN_XML, BM.YN_DOWNLOAD, BM.CD_BOARDTYPE		"
		SQL = SQL & "FROM T_BOARDMAST AS BM 																					"
		SQL = SQL & "WHERE BM.CD_BOARDCD = '" & CD_BOARDCD & "'																	"
		objADO.setSql(SQL)
		aryBoardInfo = objADO.getArrRs()

		NM_BOARDNM	= aryBoardInfo(0,0)
		YN_USE		= aryBoardInfo(1,0)
		YN_CNTVIEW	= aryBoardInfo(2,0)
		YN_ANSWER	= aryBoardInfo(3,0)
		YN_REPLY	= aryBoardInfo(4,0)
		YN_XML		= aryBoardInfo(5,0)
		YN_DOWNLOAD	= aryBoardInfo(6,0)
		CD_BOARDTYPE	= aryBoardInfo(7,0)

		SQL= ""
		SQL = SQL & "DECLARE @CNT int																					"
		SQL = SQL & "DECLARE @SNUM int																					"
		SQL = SQL & "DECLARE @ENUM int																					"
		SQL = SQL & "SET @ENUM = " & page & " * " & pgsize & "															"
		SQL = SQL & "SET @SNUM = (@ENUM - " & pgsize & ")+1																"
		SQL = SQL & "																									"
		SQL = SQL & "SELECT @CNT = COUNT(*)																				"
		SQL = SQL & "FROM                                                                                               "
		SQL = SQL & "	T_BOARDMAST AS BM                                                                               "
		SQL = SQL & "		INNER JOIN T_BOARDDETAIL AS BD                                                              "
		SQL = SQL & "		ON BD.CD_BOARDCD = BM.CD_BOARDCD															"
'		SQL = SQL & "		INNER JOIN T_ADMIN AS AM																	"
'		SQL = SQL & "		ON AM.CD_ADMINID = BD.CD_INUSER																"
		SQL = SQL & "WHERE                                                                                              "
		SQL = SQL & "	BM.CD_BOARDCD = '" & CD_BOARDCD & "' AND BD.YN_USE='Y'                                          "
		
if sKEYWORD <> "" then 
	if sGUBUN = "title" then 
		SQL = SQL & "			AND BD.NM_TITLE LIKE '%'+'" & sKEYWORD & "'+'%'													"
	elseif sGUBUN = "contents" then 
		SQL = SQL & "			AND BD.NM_CONTENTS LIKE '%'+'" & sKEYWORD & "'+'%'												"
	end if 
end if 

If sSTARTDATE <> "" Then 
		SQL = SQL & "	AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) >= '" & Int(Replace(sSTARTDATE,"-","")) & "' "
End If

If sENDDATE <> "" Then 
		SQL = SQL & "	AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) <= '" & Int(Replace(sENDDATE,"-","")) & "' "
End If

If sDISPLAY <> "" Then 
		SQL = SQL & "	AND BD.YN_DISPLAY = '" & sDISPLAY & "' "
End If

		SQL = SQL & "																											"
		SQL = SQL & "SELECT																										"
		SQL = SQL & "	@CNT AS CNT, *																							"
		SQL = SQL & "FROM																										"
		SQL = SQL & "	(																										"
		SQL = SQL & "		SELECT																								"
		SQL = SQL & "		ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]	"
		SQL = SQL & "		, BM.NM_BOARDNM, BM.CD_BOARDCD, BM.YN_USE, BM.YN_CNTVIEW, BM.YN_ANSWER, BM.YN_REPLY, BM.YN_XML		"
		'							2				3			4			5				6				7		8
		SQL = SQL & "		, BD.CD_BOARDID, BD.NO_BOARD_DEPTH, BD.NO_BOARD_SEQ, BD.YN_DISPLAY, BD.NO_VIEWCNT, BD.NO_VIEWORDER	"
		'							9				10				11				12				13				14
		SQL = SQL & "		, BD.NM_TITLE, BD.NM_CONTENTS, BD.CD_INUSER, BD.DT_INSYSDATE										"
		'							15			16				17				18				
		SQL = SQL & "		, (SELECT COUNT(*) FROM T_BOARDREPLY WHERE CD_BOARDID = BD.CD_BOARDID AND YN_USE = 'Y') AS NO_REPLY	"	'	19
		'SQL = SQL & "		, BD.NO_DOWNLOAD, BD.CD_BOARDKEY, AM.YN_LOGIN, AM.NM_NAME, BD.NM_FIELD_15							"
		'							20				21				'22			'23			'24
		SQL = SQL & "		, BD.NO_DOWNLOAD, BD.CD_BOARDKEY, BD.NM_FIELD_15							"
		'							20				21				'22
		SQL = SQL & "		,ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC) AS [NoDesc]	"	'23
		SQL = SQL & "		,BD.NM_FIELD_1							" '24
		SQL = SQL & "		FROM																								"
		SQL = SQL & "			T_BOARDMAST AS BM																				"
		SQL = SQL & "				INNER JOIN T_BOARDDETAIL AS BD																"
		SQL = SQL & "				ON BD.CD_BOARDCD = BM.CD_BOARDCD															"
'		SQL = SQL & "				INNER JOIN T_ADMIN AS AM																	"
'		SQL = SQL & "				ON AM.CD_ADMINID = BD.CD_INUSER																"
		SQL = SQL & "		WHERE                                                                                               "
		SQL = SQL & "			BM.CD_BOARDCD = '" & CD_BOARDCD & "' AND BD.YN_USE='Y'                                          "

if sKEYWORD <> "" then 
	if sGUBUN = "title" then 
		SQL = SQL & "			AND BD.NM_TITLE LIKE '%'+'" & sKEYWORD & "'+'%'													"
	elseif sGUBUN = "contents" then 
		SQL = SQL & "			AND BD.NM_CONTENTS LIKE '%'+'" & sKEYWORD & "'+'%'												"
	end if 
end if 

If sSTARTDATE <> "" Then 
		SQL = SQL & "	AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) >= '" & Int(Replace(sSTARTDATE,"-","")) & "' "
End If

If sENDDATE <> "" Then 
		SQL = SQL & "	AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) <= '" & Int(Replace(sENDDATE,"-","")) & "' "
End If

If sDISPLAY <> "" Then 
		SQL = SQL & "	AND BD.YN_DISPLAY = '" & sDISPLAY & "' "
End If

		SQL = SQL & "	) AS J																									"
		SQL = SQL & "WHERE																										"
		SQL = SQL & "	No BETWEEN @SNUM AND @ENUM																				"
		SQL = SQL & "ORDER BY																									"
		SQL = SQL & "	NO ASC																									"

		objADO.setSql(SQL)
		aryBoardList = objADO.getArrRs()

		If IsArray(aryBoardList) Then
		'If Not RS.Eof then
			tn =  aryBoardList(0,0)								'총 페이지수
			pgcount = - INT( - ( tn/pgsize ) )					'페이지수 계산
			If (page>pgcount) Then page=pgcount End If 			'페이지수 이상의 페이지를 요구하면 마지막 페이지로

			'이전page 다음page
			prevPage = page - ((page-1) Mod setsize) - setsize
			IF prevPage < 1 THEN
				prevPage = 0
			END IF
			nextPage = page - ((page-1) Mod setsize) + setsize
			IF nextPage > pgcount THEN
				nextPage = 0
			END If
		else
			tn = 0
			pgcount = 0
		End If
		
	Set objADO = Nothing

%>
<!doctype html>
<html>
<head>
	<!--#include virtual="/manager/inc/inc.ui.head.asp"-->
	<title><%=NM_BOARDNM%> 관리</title>
</head>
<body class="skin-2">	
	<div id="wrapper">

		<!--#include virtual="/manager/inc/inc.ui.navi.asp"-->

		<div id="page-wrapper" class="gray-bg dashbard-1">
			
			<!--#include virtual="/manager/inc/inc.ui.top.asp"-->
			


			<div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-10">
                    <h2><%=NM_BOARDNM%></h2>
                    <ol class="breadcrumb">
                        <li>
                            <a href="/manager/">Home</a>
                        </li>
						<!--
                        <li>
                            <a href="javascript:void(0)">게시판 관리</a>
                        </li>
						-->
                        <li class="active">
                            <strong><%=NM_BOARDNM%></strong>
                        </li>
                    </ol>
                </div>
            </div>
			
			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-12">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-search"></i> 검색 <small>게시글을 검색합니다.</small></h5>
								<div class="ibox-tools">
									<a class="collapse-link">
										<i class="fa fa-chevron-up"></i>
									</a>
								</div>
							</div>
							<div class="ibox-content m-b-sm border-bottom">
								<div class="row">
									<div class="col-sm-3">
										<div class="form-group">
											<label class="control-label" for="date_added">등록일자 - 시작</label>
											<div class="input-group date">
												<input id="sSTARTDATE" name="sSTARTDATE" type="text" class="form-control text-center" value="<%=sSTARTDATE%>" data-mask="9999-99-99" />
												<span class="input-group-addon" id="btnStartDate"><i class="fa fa-calendar"></i></span>
												<span class="input-group-addon" id="btnStartDateRemove"><i class="fa fa-refresh"></i></span>
											</div>
										</div>
									</div>
									<div class="col-sm-3">
										<div class="form-group">
											<label class="control-label" for="date_modified">등록일자 - 종료</label>
											<div class="input-group date">
												<input id="sENDDATE" name="sENDDATE" type="text" class="form-control text-center" value="<%=sENDDATE%>"  data-mask="9999-99-99" />
												<span class="input-group-addon" id="btnEndDate"><i class="fa fa-calendar"></i></span>
												<span class="input-group-addon" id="btnEndDateRemove"><i class="fa fa-refresh"></i></span>
											</div>
										</div>
									</div>
									<div class="col-sm-2">
										<div class="form-group">
											<label class="control-label" for="amount">전시여부</label>
											<select name="sDISPLAY" id="sDISPLAY" class="form-control">
												<option value="" <%=getBoolean(sDISPLAY="","selected","")%>>전체</option>
												<option value="Y" <%=getBoolean(sDISPLAY="Y","selected","")%>>예</option>
												<option value="N" <%=getBoolean(sDISPLAY="N","selected","")%>>아니오</option>
											</select>
										</div>
									</div>
								</div>

								<div class="row">
									<div class="col-sm-2">
										<div class="form-group">
											<label class="control-label" for="sGUBUN">구분</label>
											<select name="sGUBUN" id="sGUBUN" class="form-control">
												<option value="title" <%=getBoolean(sGUBUN="title","selected","")%>>제목</option>
												<option value="userid" <%=getBoolean(sGUBUN="userid","selected","")%>>아이디</option>
												<option value="usernm" <%=getBoolean(sGUBUN="usernm","selected","")%>>작성자</option>
											</select>
										</div>
									</div>
									<div class="col-sm-3">
										<div class="form-group">
											<label class="control-label" for="sKEYWORD">검색어</label>
											<div class="input-group">
												<input type="text" class="form-control" id="sKEYWORD" name="sKEYWORD" value="<%=sKEYWORD%>" placeholder="검색어를 입력해주세요." />
												<span class="input-group-btn">
													<button type="button" class="btn" onclick="fn_Search();"> <i class="fa fa-search"></i><!--조회--></button>
												</span>
											</div>
										</div>
									</div>
								</div>
							</div>
						</div>
						
						<p class="text-right">
							<button type="button" class="btn btn-sm btn-info" onclick="fn_BoardIns(<%=CD_BOARDCD%>,0,0,1,1);"><i class="fa fa-plus"></i> 등록</button>
						</p>

						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-list"></i> 리스트 <small> <strong class="text-info"><%=tn%></strong> 건의 데이터가 존재합니다.</small></h5>
							</div>
							<div class="ibox-content">
								
								<div class="table-responsive">

									<table class="table table-hover">
										<colgroup>
											<col width="5%" />
										<%
											if CD_BOARDTYPE = "GALLERY" then 
										%>
												<col width="100px" />
										<%
											end If
										%>
											<col width="" />
											<col width="150px" />
											<col width="180px" />
										<%
											if YN_CNTVIEW = "Y" then 
										%>
												<col width="10%" />
										<%
											end if

											if YN_REPLY = "Y" then 
										%>
												<col width="10%" />
										<%
											end if

											if YN_DOWNLOAD = "Y" then 
										%>
												<col width="10%" />
										<%
											end if

											if YN_ANSWER = "Y" then
										%>
												<col width="10%" />
										<%
											end if 
										%>
											<col width="100px" />
										</colgroup>
										<thead>
											<tr>
												<th class="text-center">No.</th>
										<%
											if CD_BOARDTYPE = "GALLERY" then 
												spancnt = spancnt + 1
										%>
												<th class="text-center">썸네일</th>
										<%
											end If
										%>
												<th class="text-center">제목</th>
												<th class="text-center">작성자</th>
												<th class="text-center">작성일</th>
										<%
											if YN_CNTVIEW = "Y" then 
												spancnt = spancnt + 1
										%>
												<th class="text-center">조회수</th>
										<%
											end if

											if YN_REPLY = "Y" then 
												spancnt = spancnt + 1
										%>
												<th class="text-center">덧글수</th>
										<%
											end if

											if YN_DOWNLOAD = "Y" then 
												spancnt = spancnt + 1
										%>
												<th class="text-center">다운로드</th>
										<%
											end if

											if YN_ANSWER = "Y" then
												spancnt = spancnt + 1
										%>
												<th class="text-center">답글</th>

										<%
											end if 
										%>
												<th class="text-center">전시여부</th>
											</tr>
										</thead>
										<tbody>
										<%
											If IsArray(aryBoardList) Then
											For iloop = LBound(aryBoardList,2) To UBound(aryBoardList,2)
										%>
											<tr>
												<td class="text-center"><%=aryBoardList(23,iloop)%><%'=aryBoardList(1,iloop)%></td>
												<%if CD_BOARDTYPE = "GALLERY" then %>
												<td class="text-center">
													<%if aryBoardList(24,iloop) <> "" then %>
														<img src="/file/board/<%=aryBoardList(24,iloop)%>" style="max-width: 100px; max-height: 50px" onerror='' />
													<%End If %>		
												</td>
												<%end If%>
												<td class="text-left">
													<a href="javascript:fn_boardview('<%=aryBoardList(9,iloop)%>');">
														<%
														if aryBoardList(10,iloop) > 1 then 
															for jloop = 2 to aryBoardList(10,iloop)
															response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
															next 
															response.write "-> RE"
														end if 
														response.write "&nbsp;&nbsp;" & aryBoardList(15,iloop)
														%>
													</a>
												</td>
												<td class="text-center">
													<%=aryBoardList(17,iloop)%>
												</td>
												<td class="text-center"><%=aryBoardList(18,iloop)%></td>
										<%if YN_CNTVIEW = "Y" then %>
												<td class="text-center"><%=aryBoardList(13,iloop)%></td>
										<%end if %>
										<%if YN_REPLY = "Y" then %>
												<td class="text-center"><%=aryBoardList(19,iloop)%></td>
										<%end if %>
										<%if YN_DOWNLOAD = "Y" then %>
												<td class="text-center"><%=aryBoardList(20,iloop)%></td>
										<%end if %>
										<%if YN_ANSWER = "Y" then%>
												<td class="text-center">
													<button type="button" class="btn" onclick="fn_BoardIns(<%=CD_BOARDCD%>,<%=aryBoardList(9,iloop)%>,<%=aryBoardList(21,iloop)%>,<%=aryBoardList(10,iloop)+1%>,<%=aryBoardList(11,iloop)+1%>);"><i class="fa fa-search"></i> 답글</button>
												</td>
										<%end if %>
												<td class="text-center">
													<select class="form-control text-center" id="YN_DISPLAY" name="YN_DISPLAY" onChange="fn_changeDisplay('<%=aryBoardList(9,iloop)%>',this.value);">
														<option value="Y" <%=getBoolean(aryBoardList(12,iloop)="Y","selected","")%>>예</option>
														<option value="N" <%=getBoolean(aryBoardList(12,iloop)="N","selected","")%>>아니오</option>
													</select>
												</td>
											</tr>
										<%
											Next
										Else 
										%>
											<tr>
												<td class="text-center" colspan="<%=spancnt%>">데이터가 존재하지 않습니다.</td>
											</tr>
										<%
										End If 
										%>
										</tbody>
										<tfoot>
											<tr>
												<td class="text-center" colspan="<%=spancnt%>">
													<%
														Call PageNavigation("/manager/board/BoardDetail.asp","CD_BOARDCD="&CD_BOARDCD&"&sGUBUN="&sGUBUN&"&sKEYWORD="&sKEYWORD&"&sDISPLAY="&sDISPLAY&"&sSTARTDATE="&sSTARTDATE&"&sENDDATE="&sENDDATE, tn, pgcount, startpage, page, setsize, prevpage, nextpage)
													%>
												</td>
											</tr>
										</tfoot>
									</table>
								</div>

							</div>
						</div>

					</div>
				</div>
			</div>

			<!--#include virtual="/manager/inc/inc.ui.footer.asp"-->
		</div>
	</div>

	<!--#include virtual="/manager/inc/inc.ui.hidden.asp"-->

	<script type="text/javascript">
		$(document).ready(function() {
		
            $('#sSTARTDATE').datepicker({
                todayBtn: "linked",
                keyboardNavigation: false,
                forceParse: false,
                autoclose: true,
				format: 'yyyy-mm-dd',
				keyboardNavigation: true,
				language: 'ko'
            });

            $('#sENDDATE').datepicker({
                todayBtn: "linked",
                keyboardNavigation: false,
                forceParse: false,
                autoclose: true,
				format: 'yyyy-mm-dd',
				keyboardNavigation: true,
				language: 'ko'
            });
			
			$('#btnStartDate').click(function() {
				$('#sSTARTDATE').focus();
			});

			$('#btnStartDateRemove').click(function() {
				$('#sSTARTDATE').val("");
			});

			$('#btnEndDate').click(function() {
				$('#sENDDATE').focus();
			});

			$('#btnEndDateRemove').click(function() {
				$('#sENDDATE').val("");
			});
        });

		function fn_Search()
		{
			var sgubun = $("#sGUBUN option:selected").val();
			var skeyword = $("#sKEYWORD").val();
			var sDisplay = $("#sDISPLAY option:selected").val();
			var sStartdate = $("#sSTARTDATE").val();
			var sEnddate = $("#sENDDATE").val();

			document.location.href = "BoardDetail.asp?CD_BOARDCD=<%=CD_BOARDCD%>&sGUBUN="+sgubun+"&sKEYWORD="+skeyword+"&sDISPLAY="+sDisplay+"&sSTARTDATE="+sStartdate+"&sENDDATE="+sEnddate;
		}

		function fn_boardview(CD_BOARDID)
		{
			document.location.href="BoardDetailView.asp?page=<%=page%>&CD_BOARDID="+CD_BOARDID;
		}

		function fn_BoardIns(CD_BOARDCD,CD_BOARDID,CD_BOARDKEY,NO_BOARD_DEPTH,NO_BOARD_SEQ)
		{
			document.location.href="BoardDetailIns.asp?CD_BOARDCD="+CD_BOARDCD+"&CD_BOARDID="+CD_BOARDID+"&CD_BOARDKEY="+CD_BOARDKEY+"&NO_BOARD_DEPTH="+NO_BOARD_DEPTH+"&NO_BOARD_SEQ="+NO_BOARD_SEQ;
		}

		function fn_read(CD_BOARDID)
		{
			window.open("./BoardRead.asp?CD_BOARDID="+CD_BOARDID);
		}

		function fn_changeDisplay(CD_BOARDID, YN_DISPLAY)
		{
			var hd_frame = document.getElementById("hd_frame");
			hd_frame.src = "BoardDetail_display_proc.asp?CD_BOARDID=" + CD_BOARDID + "&YN_DISPLAY=" + YN_DISPLAY;
		}
	</script>

	<iframe type="hiddenframe" id="hd_frame" name="hd_frame" style="display:none;"></iframe>
</body>
</html>		