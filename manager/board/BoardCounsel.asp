<!--#include virtual="/Include/Config.asp"-->
<!--#include virtual="/function/fn_AdminLoginCheck.asp"-->
<!--#include virtual="/function/fn_PageNavigation.asp"-->
<!--#include virtual="/function/fn_CodeInfo.asp"-->
<%
	'########################################################################################
	'#	File		: 
	'#  Create		: 
	'#	Info		: 
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	Dim CD_BOARDCD : CD_BOARDCD = ReqQ("CD_BOARDCD")
	Dim sSTARTDATE : sSTARTDATE = ReqQ("sSTARTDATE")	'기간검색(시작일)
	Dim sENDDATE : sENDDATE = ReqQ("sENDDATE")	'기간검색(종료일)

	Dim sGUBUN : sGUBUN = ReqQ("sGUBUN")
	Dim sKEYWORD : sKEYWORD = ReqQ("sKEYWORD")
	Dim sSTATUS : sSTATUS = ReqQ("sSTATUS") '답변 처리 상태

	Dim objADO, SQL, iloop, jloop, kloop 
	Dim SQL_where, SQL_device_type, SQL_media_type, SQL_counsel_ratio
	Dim aryChartDeviceType, aryChartMediaType, aryChartCounselRatioType
	Dim aryBoardList, aryBoardInfo, aryCodeInfo, arySTATUS, aryCOUNSEL_TYPE
	Dim NM_BOARDNM, YN_USE, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, YN_DOWNLOAD
	Dim spancnt : spancnt = 12

	If CD_BOARDCD = "1003" Then 
		'온라인상담
		spancnt = 11
	ElseIf CD_BOARDCD = "1007" Then 
		'카카오톡상담
		spancnt = 8
	ElseIf CD_BOARDCD = "1008" Then 
		'빠른 상담
		spancnt = 8
	End If 
	
	If CD_BOARDCD = "" Then
		Response.Write "<script>alert('올바른경로로 접근하세요.');history.back();</script>"
		Response.End
	End If
	
	' init(답변 처리상태 검색)
	If sSTATUS = "" Then 
		sSTATUS = "ALL"	'답변 준비중부터 가져오기(ALL/READY/SUCCESS)
	End If
	
	' init(검색 - 처리상태 리스트 가져오기)
	arySTATUS = fn_CodeInfo("STATUS")
	
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
		
		'// start: 검색 조건들 -------------------------------------------- 
		SQL_where = ""
		' 키워드 검색 조건
		If sKEYWORD <> "" Then
			If sGUBUN = "usernm" Then 
				SQL_where = SQL_where & "			AND BD.NM_FIELD_1 LIKE '" & sKEYWORD & "%'														"
			End If 
		End If
			
		' 기간 검색 조건
		If sSTARTDATE <> "" Then 
				SQL_where = SQL_where & "			AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) >= " & Replace(sSTARTDATE,"-","") & "		"
		End If 

		If sENDDATE <> "" Then 
				SQL_where = SQL_where & "			AND CONVERT(VARCHAR(8), BD.DT_INSYSDATE, 112) <= " & Replace(sENDDATE,"-","") & "		"
		End If 

		' 답변 처리상태 조건
		If sSTATUS <> "" And sSTATUS <> "ALL" Then 
				SQL_where = SQL_where & "			AND BD.NM_FIELD_11 = '" & sSTATUS & "'													"
		End If
		'// end: 검색 조건들 -------------------------------------------- 

		'// start: 상담 게시글 조회 -------------------------------------------- 
		SQL = ""
		SQL = SQL & "SELECT BM.NM_BOARDNM, BM.YN_USE, BM.YN_CNTVIEW, BM.YN_ANSWER, BM.YN_REPLY, BM.YN_XML, BM.YN_DOWNLOAD	"
		SQL = SQL & "FROM T_BOARDMAST AS BM 																							"
		SQL = SQL & "WHERE BM.CD_BOARDCD = '" & CD_BOARDCD & "'																			"
		objADO.setSql(SQL)
		aryBoardInfo = objADO.getArrRs()

		NM_BOARDNM	= aryBoardInfo(0,0)
		YN_USE		= aryBoardInfo(1,0)
		YN_CNTVIEW	= aryBoardInfo(2,0)
		YN_ANSWER	= aryBoardInfo(3,0)
		YN_REPLY	= aryBoardInfo(4,0)
		YN_XML		= aryBoardInfo(5,0)
		YN_DOWNLOAD	= aryBoardInfo(6,0)
		
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
		SQL = SQL & "		INNER JOIN T_ADMIN AS AM																	"
		SQL = SQL & "		ON AM.CD_ADMINID = BD.CD_INUSER																"
		
		SQL = SQL & "		INNER JOIN T_CODE AS CD2																	"
		SQL = SQL & "		ON CD2.CD_CODE = BD.NM_FIELD_11	AND CD2.CD_GUBUN = 'STATUS' AND CD2.YN_USE = 'Y'			"
		SQL = SQL & "		WHERE                                                                                       "
		SQL = SQL & "			BM.CD_BOARDCD = '" & CD_BOARDCD & "'  "
		SQL = SQL & "															"
		SQL = SQL & "			AND BD.YN_USE = 'Y' AND BD.NO_BOARD_DEPTH = 1	"
		
		SQL = SQL & SQL_where

		SQL = SQL & "																											"
		SQL = SQL & "SELECT																										"
		SQL = SQL & "	@CNT AS CNT, *																							"
		SQL = SQL & "FROM																										"
		SQL = SQL & "	(																										"
		SQL = SQL & "		SELECT																								"
		'SQL = SQL & "		, CD2.CD_CODE AS [CD2.CD_CODE], CD2.NM_CODE AS [CD2.NM_CODE], BD.NM_FIELD_3, BD.NM_FIELD_6, BD.NM_FIE

		SQL = SQL & "			ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY DESC, BD.NO_BOARD_SEQ ASC) AS [No]		"	'1
		SQL = SQL & "			,ROW_NUMBER() OVER (ORDER BY BD.NO_VIEWORDER ASC, BD.CD_BOARDKEY ASC, BD.NO_BOARD_SEQ ASC) AS [NoDesc]	"	'2
		SQL = SQL & "			,BM.NM_BOARDNM,BD.CD_BOARDID,BD.CD_BOARDCD,BD.YN_DISPLAY,BD.NO_VIEWCNT									"	'7
		SQL = SQL & "			,BD.YN_USE,BD.NM_TITLE,BD.NM_CONTENTS																	"	'10
		SQL = SQL & "			,BD.NM_FIELD_1,BD.NM_FIELD_2,BD.NM_FIELD_3,BD.NM_FIELD_4,BD.NM_FIELD_5									"	'15
		SQL = SQL & "			,BD.NM_FIELD_6,BD.NM_FIELD_7,BD.NM_FIELD_8,BD.NM_FIELD_9,BD.NM_FIELD_10									"	'20
		SQL = SQL & "			,BD.NM_FIELD_11,BD.NM_FIELD_12,BD.NM_FIELD_13,BD.NM_FIELD_14,BD.NM_FIELD_15								"	'25
		SQL = SQL & "			,BD.CD_INUSER,BD.DT_INSYSDATE																			"	'27

		SQL = SQL & "			, BD.NO_INIPADDR																						"	'28
		SQL = SQL & "			, BD.HTTP_REFERER, BD.HTTP_USER_AGENT, BD.HTTP_DEVICE_TYPE, BD.HTTP_MEDIA_TYPE							"	'32
		SQL = SQL & "			, BD.NO_BOARD_DEPTH							"	'33

		SQL = SQL & "			, (SELECT TOP 1 ANS.NM_CONTENTS FROM T_BOARDDETAIL AS ANS WHERE ANS.CD_BOARDCD = '" & CD_BOARDCD & "' AND ANS.NO_BOARD_DEPTH = 2 AND ANS.CD_BOARDKEY = BD.CD_BOARDKEY ORDER BY ANS.CD_BOARDID DESC) AS ANSWER_CONTENTS	"	'34
		SQL = SQL & "			, (SELECT TOP 1 ANS.CD_BOARDID FROM T_BOARDDETAIL AS ANS WHERE ANS.CD_BOARDCD = '" & CD_BOARDCD & "' AND ANS.NO_BOARD_DEPTH = 2 AND ANS.CD_BOARDKEY = BD.CD_BOARDKEY ORDER BY ANS.CD_BOARDID DESC) AS ANSWER_BOARDID	"	'35
		

		SQL = SQL & "		FROM																								"
		SQL = SQL & "			T_BOARDMAST AS BM																				"
		SQL = SQL & "				INNER JOIN T_BOARDDETAIL AS BD																"
		SQL = SQL & "				ON BD.CD_BOARDCD = BM.CD_BOARDCD															"
		SQL = SQL & "				INNER JOIN T_ADMIN AS AM																	"
		SQL = SQL & "				ON AM.CD_ADMINID = BD.CD_INUSER																"
		
		SQL = SQL & "				INNER JOIN T_CODE AS CD2																	"
		SQL = SQL & "				ON CD2.CD_CODE = BD.NM_FIELD_11	AND CD2.CD_GUBUN = 'STATUS' AND CD2.YN_USE = 'Y' "
		SQL = SQL & "		WHERE                                                                                               "
		SQL = SQL & "				BM.CD_BOARDCD = '" & CD_BOARDCD & "' "
		SQL = SQL & "																	"
		SQL = SQL & "			AND BD.YN_USE = 'Y' AND BD.NO_BOARD_DEPTH = 1	"

		SQL = SQL & SQL_where

		SQL = SQL & "	) AS J																									"
		SQL = SQL & "WHERE																										"
		SQL = SQL & "	No BETWEEN @SNUM AND @ENUM																				"
		SQL = SQL & "ORDER BY																									"
		SQL = SQL & "	NO ASC																									"
		
		'Response.write SQL

		objADO.setSql(SQL)
		aryBoardList = objADO.getArrRs()

		If IsArray(aryBoardList) Then
		'If Not RS.Eof then
			tn =  aryBoardList(0,0)								'총 페이지수
			pgcount = - INT( - ( tn/pgsize ) )					'페이지수 계산
			If (page>pgcount) Then page=pgcount End If 			'페이지수 이상의 페이지를 요구하면 마지막 페이지로

			'이전page 다음page
			prevPage = page - ((page-1) Mod setsize) - setsize
			If prevPage < 1 Then 
				prevPage = 0
			End If 
			nextPage = page - ((page-1) Mod setsize) + setsize
			If nextPage > pgcount Then 
				nextPage = 0
			End If
		Else 
			tn = 0
			pgcount = 0
		End If
		
		'// end: 상담 게시글 조회 -------------------------------------------- 

	Set objADO = Nothing

	' 상담구분
	Dim aryCounselType : aryCounselType = fn_CodeInfo("COUNSEL_TYPE")
%>
<!doctype html>
<html>
<head>
	<!--#include virtual="/manager/inc/inc.ui.head.asp"-->
	<title><%=NM_BOARDNM%> 관리</title>
	<style type="text/css">
		.custom-contents p {margin: 0 !important;}
	</style>
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
                            <a href="javascript:void(0)">상담</a>
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
						
						<div class="ibox float-e-margins"><!-- collapsed-->
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
											<label class="control-label small" for="date_added">등록일자 - 시작</label>
											<div class="input-group date">
												<input id="sSTARTDATE" name="sSTARTDATE" type="text" class="form-control text-center" value="<%=sSTARTDATE%>" data-mask="9999-99-99" />
												<span class="input-group-addon" id="btnStartDate"><i class="fa fa-calendar"></i></span>
												<span class="input-group-addon" id="btnStartDateRemove"><i class="fa fa-refresh"></i></span>
											</div>
										</div>
									</div>
									<div class="col-sm-3">
										<div class="form-group">
											<label class="control-label small" for="date_modified">등록일자 - 종료</label>
											<div class="input-group date">
												<input id="sENDDATE" name="sENDDATE" type="text" class="form-control text-center" value="<%=sENDDATE%>"  data-mask="9999-99-99" />
												<span class="input-group-addon" id="btnEndDate"><i class="fa fa-calendar"></i></span>
												<span class="input-group-addon" id="btnEndDateRemove"><i class="fa fa-refresh"></i></span>
											</div>
										</div>
									</div>
									<div class="col-sm-2">
										<div class="form-group">
											<label class="control-label small" for="amount">처리상태</label>
											<select name="sSTATUS" id="sSTATUS" class="form-control">
												<option value="ALL" <%=getBoolean(sSTATUS="ALL","selected","")%>>전체</option>
											<%
											If IsArray(arySTATUS) Then 
												For iloop = 0 To UBound(arySTATUS,2)
											%>
												<option value="<%=arySTATUS(1,iloop)%>" <%=getBoolean(sSTATUS=arySTATUS(1,iloop),"selected","")%> style="color: <%=arySTATUS(9,iloop)%>"><%=arySTATUS(2,iloop)%></option>
											<%
												Next 
											End If 
											%>
											</select>
										</div>
									</div>
								</div>

								<div class="row">
									<div class="col-sm-2">
										<div class="form-group">
											<label class="control-label small" for="sGUBUN">구분</label>
											<select name="sGUBUN" id="sGUBUN" class="form-control">
												<!--
												<option value="title" <%=getBoolean(sGUBUN="title","selected","")%>>제목</option>
												<option value="userid" <%=getBoolean(sGUBUN="userid","selected","")%>>아이디</option>
												-->
												<option value="usernm" <%=getBoolean(sGUBUN="usernm","selected","")%>>작성자</option>
											</select>
										</div>
									</div>
									<div class="col-sm-3">
										<div class="form-group">
											<label class="control-label small" for="sKEYWORD">검색어</label>
											<div class="input-group">
												<input type="text" class="form-control" id="sKEYWORD" name="sKEYWORD" value="<%=sKEYWORD%>" placeholder="검색어를 입력해주세요." onkeypress="fn_chkent();" />
												<span class="input-group-btn">
													<button type="button" class="btn" onclick="fn_Search();"> <i class="fa fa-search"></i><!--조회--></button>
												</span>
											</div>
										</div>
									</div>
								</div>
							</div>
						</div>
						

						<form id="frm_BoardDetail" name="frm_BoardDetail">
	
							<input type="hidden" id="COUNSEL_GUBUN" name="COUNSEL_GUBUN" value="" /> <!--처리결과(state)/메모(memo) 구분-->
							<input type="hidden" id="CD_BOARDID" name="CD_BOARDID" value="" />		 <!--처리결과/메모 수정 게시글 아이디-->

							<div class="ibox float-e-margins">
								<div class="ibox-title">
									<h5><i class="fa fa-list"></i> 리스트 <small> <strong class="text-info"><%=tn%></strong> 건의 데이터가 존재합니다.</small></h5>
									<div class="ibox-tools">
									<!--
										<a class="collapse-link">
											<i class="fa fa-chevron-up"></i>
										</a>
										<a class="dropdown-toggle" data-toggle="dropdown" href="#">
											<i class="fa fa-wrench"></i>
										</a>
									-->
										<!--
										<a class="close-link">
											<i class="fa fa-times"></i>
										</a>
										-->
									</div>
								</div>
								<div class="ibox-content">
									
									<div class="table-responsive">

										<table class="table table-striped table-bordered">
											<colgroup>
												<col width="50px" />
												<col width="30px" />
												<col width="70px" />
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>
												<col width="110px" />
												<%End If %>
												<col width="60px" />
												<col width="80px" />
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>
												<col width="120px" />
												<%End If %>
												<%If CD_BOARDCD = "1007" Or CD_BOARDCD = "1008" Then	'카카오톡, 빠른상담%>
												<col width="120px" />
												<%End If %>
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>
												<col width="200px" />
												<%End If %>
												<col width="100px" />
												<col width="200px" />
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>
												<col width="70px" />
												<%End If %>
											</colgroup>
											<thead>
												<tr>
													<th class="text-center small">No</th>
													<th class="text-center small">PC/MOBILE</th>
													<th class="text-center small">등록일</th>
													<%If CD_BOARDCD = "1003" Then	'온라인상담%>
													<th class="text-center small"><!--상담과목-->상태</th>
													<%End If %>
													<th class="text-center small">작성자</th>
													<th class="text-center small">핸드폰번호</th>
													<%If CD_BOARDCD = "1003" Then	'온라인상담%>
													<th class="text-center small">이메일</th>
													<%End If %>
													<%If CD_BOARDCD = "1007" Or CD_BOARDCD = "1008" Then	'카카오톡, 빠른상담%>
													<th class="text-center small">문의내용</th>
													<%End If %>
													<%If CD_BOARDCD = "1003" Then	'온라인상담%>
													<th class="text-center small">상담내용</th>
													<%End If %>
													<th class="text-center small">처리상태</th>
													<th class="text-center small">메모</th>
													<%If CD_BOARDCD = "1003" Then	'온라인상담%>
													<th class="text-center small">답변</th>
													<%End If %>
												</tr>
											</thead>
											<tbody>
											<%
											If IsArray(aryBoardList) Then 
												For iloop = 0 To UBound(aryBoardList,2)
											%>
												<tr>
													<td class="text-center small"><%=aryBoardList(2,iloop)%></td>
													<td class="text-center small">
														<%
														If aryBoardList(31,iloop) = "DESKTOP" Then 
															Response.write "PC"
														ElseIf aryBoardList(31,iloop) = "MOBILE" Then 
															Response.write "MOBILE"
														Else
															Response.write "PC"
														End If
														%>
													</td>
													
													<td class="text-center small">
														<%=Left(aryBoardList(27,iloop), 10)%>
														<br />
														<%=Right(aryBoardList(27,iloop), (Len(aryBoardList(27,iloop)) - 11))%>
													</td>
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>
													<td class="text-center small">
														<%
														For jloop = 0 To UBound(aryCounselType, 2)
															If aryCounselType(1,jloop) = aryBoardList(14,iloop) Then 
														%>
															<%=aryCounselType(2,jloop)%>
														<%
															End If 
														Next
														%>
													</td>
												<%End If %>
													<td class="text-center small"><a href="javascript:fn_boardview('<%=aryBoardList(4,iloop)%>');"><%=aryBoardList(11,iloop)%></a></td>
													<td class="text-center small">
														<%=addHyphen(aryBoardList(12,iloop))%>
														<!--<br />[답변수신여부: <strong><%=aryBoardList(18,iloop)%></strong>]-->
													</td>
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>	
													<td class="text-center small">
														<%=aryBoardList(13,iloop)%>
														<%If aryBoardList(13,iloop) <> "" And aryBoardList(19,iloop) <> "" Then %>
														<br />[답변수신여부: <strong><%=aryBoardList(19,iloop)%></strong>]
														<%End If%>
													</td>
												<%End If %>
												<%If CD_BOARDCD = "1007" Or CD_BOARDCD = "1008" Then	'카카오톡, 빠른상담%>
													<td class="text-center small">
														<textarea class="form-control custom-contents no-padding" style="width: 100%; height: 120px; resize: none; background: #fff;" readonly="readonly"><%=getHtml(aryBoardList(10,iloop))%></textarea>
													</td>
												<%End If %>
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>
													<td class="no-padding">
														<textarea class="form-control custom-contents no-padding" style="width: 100%; height: 120px; resize: none; background: #fff;" readonly="readonly"><%=getHtml(aryBoardList(10,iloop))%></textarea>
														<div class="text-right">
														<%If aryBoardList(15,iloop) <> "" Then%>
															<a class="btn btn-xs btn-success btn-outline" href="/file/board/<%=aryBoardList(15,iloop)%>" target="_blank"><i class="fa fa-photo"></i> 사진1</a>
														<%End If%>
														<%If aryBoardList(16,iloop) <> "" Then%>
															<a class="btn btn-xs btn-success btn-outline" href="/file/board/<%=aryBoardList(16,iloop)%>" target="_blank"><i class="fa fa-photo"></i> 사진2</a>
														<%End If%>
														</div>
													</td>
												<%End If %>
													<td class="text-center">
														<select class="form-control" name="NM_STATE_<%=aryBoardList(4,iloop)%>" onChange="fn_counsel('state','<%=aryBoardList(4,iloop)%>');" style="font-size: 85%;">
														<%
														If IsArray(arySTATUS) Then 
															For jloop = 0 To UBound(arySTATUS,2)
														%>
															<option value="<%=arySTATUS(1,jloop)%>" <%=getBoolean(aryBoardList(21,iloop)=arySTATUS(1,jloop),"selected","")%> style="color: <%=arySTATUS(9,jloop)%>"><%=arySTATUS(2,jloop)%></option>
														<%
															Next 
														End If 
														%>
														</select>
													</td>
													<td class="text-center no-padding">
														<textarea name="NM_MEMO_<%=aryBoardList(4,iloop)%>" class="form-control" style="width: 100%; height: 120px; resize: none; font-size: 85%;"><%=getHtml(aryBoardList(22,iloop))%></textarea>
														<div class="text-right">
															<button type="button" class="btn btn-xs btn-info" onclick="fn_counsel('memo','<%=aryBoardList(4,iloop)%>');"><i class="fa fa-check"></i> 메모저장</button>
														</div>
													</td>
												<%If CD_BOARDCD = "1003" Then	'온라인상담%>
													<td class="text-center no-padding">
														<input type="hidden" id="hidden_NM_TITLE_<%=aryBoardList(4,iloop)%>" value="<%=aryBoardList(9,iloop)%>" style="display: none;" />
														<textarea id="hidden_NM_CONTENTS_<%=aryBoardList(4,iloop)%>" style="display: none;"><%=getText(aryBoardList(10,iloop))%></textarea>
														<textarea id="hidden_NM_ANSWER_<%=aryBoardList(4,iloop)%>" style="display: none;"><%=getHtml(aryBoardList(34,iloop))%></textarea>
														<button type="button" class="btn btn-xs btn-success" onclick="fn_answer('<%=aryBoardList(4,iloop)%>','<%=aryBoardList(35,iloop)%>')">
															<i class="fa fa-edit"></i> 답변작성
														</button>
													</td>
												<%End If %>
												</tr>
												
											<%
												Next
											Else
											%>
												<tr>
													<td colspan="<%=spancnt%>" class="text-center">데이터가 없습니다.</td>
												</tr>
											<%
											End If 
											%>							

											</tbody>
											<tfoot>
												<tr>
													<td colspan="<%=spancnt%>" class="text-center">
														<%
															Call PageNavigation("/manager/board/BoardCounsel.asp","CD_BOARDCD="&CD_BOARDCD&"&sGUBUN="&sGUBUN&"&sKEYWORD="&sKEYWORD&"&sSTARTDATE="&sSTARTDATE&"&sENDDATE="&sENDDATE&"&sSTATUS="&sSTATUS, tn, pgcount, startpage, page, setsize, prevpage, nextpage)
														%>
													</td>
												</tr>
											</tfoot>
										</table>
									</div>
								</div>
							</div>
						</form>

					</div>
				</div>
			</div>
			
			<!--#include virtual="/manager/inc/inc.ui.footer.asp"-->
		</div>
	</div>

	<!--#include virtual="/manager/inc/inc.ui.hidden.asp"-->


<script type="text/javascript"> 
	
	// 달력 컨트롤 가져오기
	$(document).ready(function(){
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

		$('#btnToggleBoardCounselGPSInfo').click(function() {
			if ($(this).data('openstatus') == 'open') {
				$('.board-counsel-gps-info').fadeOut('fast');
				$(this).data('openstatus', 'close');
				$(this).text('신청 GPS 좌표 열기');
			} else {
				$('.board-counsel-gps-info').fadeIn('fast');
				$(this).data('openstatus', 'open');
				$(this).text('신청 GPS 좌표 닫기');
			}
			
		});
	});
	
	/**
	 * name : fn_Search
	 * parameter : 없음
	 * description : 조회
	 */
	function fn_Search()
	{
		var sgubun = document.getElementById("sGUBUN").value;
		var skeyword = document.getElementById("sKEYWORD").value;
		var sSTARTDATE = document.getElementById("sSTARTDATE");
		var sENDDATE = document.getElementById("sENDDATE");
		var sSTATUS = document.getElementById("sSTATUS");

		document.location.href = "BoardCounsel.asp?CD_BOARDCD=<%=CD_BOARDCD%>&sGUBUN="+sgubun+"&sKEYWORD="+skeyword+"&sSTARTDATE="+sSTARTDATE.value+"&sENDDATE="+sENDDATE.value+"&sSTATUS="+sSTATUS.value;
	}

	/**
	 * name : fn_boardview
	 * parameter : CD_BOARDID
	 * description : 뷰페이지 보기
	 */
	//function fn_boardview(CD_BOARDID)
	//{
	//	document.location.href="BoardCounselView.asp?page=<%=page%>&CD_BOARDID="+CD_BOARDID;
	//}
	function fn_boardview(CD_BOARDID)
	{
		document.location.href="BoardDetailView.asp?page=<%=page%>&CD_BOARDID="+CD_BOARDID;
	}

	/**
	 * name : fn_BoardIns
	 * parameter : CD_BOARDCD,CD_BOARDID,CD_BOARDKEY,NO_BOARD_DEPTH,NO_BOARD_SEQ
	 * description : 새글 작성
	 */
	function fn_BoardIns(CD_BOARDCD,CD_BOARDID,CD_BOARDKEY,NO_BOARD_DEPTH,NO_BOARD_SEQ)
	{
		document.location.href="BoardCounselIns.asp?CD_BOARDCD="+CD_BOARDCD+"&CD_BOARDID="+CD_BOARDID+"&CD_BOARDKEY="+CD_BOARDKEY+"&NO_BOARD_DEPTH="+NO_BOARD_DEPTH+"&NO_BOARD_SEQ="+NO_BOARD_SEQ;
	}

	/**
	 * name : fn_read
	 * parameter : CD_BOARDID
	 * description : 읽기 페이지
	 */
	function fn_read(CD_BOARDID)
	{
		window.open("./BoardRead.asp?CD_BOARDID="+CD_BOARDID);
	}

	/**
	 * name : fn_changeDisplay
	 * parameter : CD_BOARDID, YN_DISPLAY
	 * description : 전시여부 변경
	 */
	function fn_changeDisplay(CD_BOARDID, YN_DISPLAY)
	{
		var hd_frame = document.getElementById("hd_frame");
		hd_frame.src = "BoardDetail_display_proc.asp?CD_BOARDID=" + CD_BOARDID + "&YN_DISPLAY=" + YN_DISPLAY;
	}
	
	/**
	 * name : fn_addZero
	 * parameter : int형 숫자
	 * description : 숫자앞 '0' 추가	
	 */
	function fn_addZero(val)
	{
		if (val < 10)
		{
			return "0" + val;
		}
		else
		{
			return "" + val;
		}
	}

	/**
	 * name : fn_chkent
	 * parameter : 없음
	 * description : 엔터키 입력시 액션
	 */
	function fn_chkent()
	{
		if (event.keyCode == 13)
		{
			fn_Search();
		}
	}
	
	/**
	 *	처리결과, 메모 등록/수정
	 */
	function fn_counsel(gubun, BOARDID)
	{
		var COUNSEL_GUBUN = document.getElementById("COUNSEL_GUBUN");
		var CD_BOARDID = document.getElementById("CD_BOARDID");

		COUNSEL_GUBUN.value = gubun;
		CD_BOARDID.value = BOARDID;

		document.frm_BoardDetail.target = "hd_frame";
		document.frm_BoardDetail.method = "post";
		document.frm_BoardDetail.action = "BoardCounsel_memo_proc.asp";
		document.frm_BoardDetail.submit();
	}

	function fn_link(url) {
		if (url != "") {
			window.open(url,"","");
		} else {
			alert("URI를 찾을 수 없습니다.");
		}
	}

	function fn_answer(BOARDID, ANSWER_BOARDID) {
		$("#MODE").val("");
		$("#question_CD_BOARDID").val("");
		$("#question_title").html("");
		$("#question_contents").html("");
		$("#answer_contents").val("");
		$("#answer_CD_BOARDID").val("");
		
		if (ANSWER_BOARDID == "") {
			$("#MODE").val("I");
		} else {
			$("#MODE").val("U");
		}

		$("#question_CD_BOARDID").val(BOARDID);
		$("#question_title").html($("#hidden_NM_TITLE_" + BOARDID).val());
		$("#question_contents").html($("#hidden_NM_CONTENTS_" + BOARDID).val());
		$("#answer_contents").val($("#hidden_NM_ANSWER_" + BOARDID).val());
		$("#answer_CD_BOARDID").val(ANSWER_BOARDID);

		$("#answerModal").modal();	
	}

	function fn_answer_submit() {

		if ($.trim($("#answer_contents").val()) == "") {
			alert("답변 내용을 입력해주세요.");
			$("#answer_contents").focus();
			return;
		}

		$("#frmAnswer").attr({
			"method": "post",
			"enctype": "multipart/form-data",
			"target": "hd_frame",
			"action": "BoardCounsel_Answer_proc.asp"
		}).submit();
	}
</script>

<div class="modal inmodal" id="answerModal" tabindex="-1" role="dialog" aria-hidden="true">
	<div class="modal-dialog">
		<div class="modal-content animated bounceInRight">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">&times;</span><span class="sr-only">Close</span></button>
				<i class="fa fa-laptop modal-icon"></i>
				<h5 class="modal-title">온라인상담 답변 작성</h5>
				<!--<small class="font-bold">Lorem Ipsum is simply dummy text of the printing and typesetting industry.</small>-->
			</div>
			<div class="modal-body">
				<div class="form-group">
					<label>제목</label>
					<p id="question_title">
						<!--제목-->
					</p>
				</div>
				<div class="form-group">
					<label>상담내용</label>
					<p id="question_contents">
						<!--문의내용-->
					</p>
				</div>
				<div class="form-group">
					<form id="frmAnswer">
						<input type="hidden" id="question_CD_BOARDID" name="question_CD_BOARDID" value="" />
						<input type="hidden" id="MODE" name="MODE" value="" />
						<input type="hidden" id="answer_CD_BOARDID" name="answer_CD_BOARDID" value="" />
						<label>답변 작성</label>
						<textarea type="text" id="answer_contents" name="answer_contents" placeholder="답변을 등록해주세요." class="form-control" rows="7" style="resize: none;"></textarea>
					</form>
				</div>
			</div>
			<div class="modal-footer">
				<button type="button" class="btn btn-white" data-dismiss="modal">닫기</button>
				<button type="button" class="btn btn-primary" onclick="fn_answer_submit()">저장</button>
			</div>
		</div>
	</div>
</div>


<iframe type="hiddenframe" id="hd_frame" name="hd_frame" style="display:none;"></iframe>
</body>
</html>