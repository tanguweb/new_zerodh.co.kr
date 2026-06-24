<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_PageNavigation.asp"-->
<%
	'########################################################################################
	'#	File		: /manager/board/BoardMast.asp
	'#  Create		: / 2010.09.06
	'#	Info		: 게시판 마스터 관리
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, SQL, AryList, AryBrandId, AryDataType, AryBoardType, AryRequireType, iloop, jloop, kloop
	Dim sYN_USE, sKEYWORD

	sYN_USE = Req("sYN_USE")
	sKEYWORD = Req("sKEYWORD")

	Dim page, startpage, tn, pgcount, nextpage, prevpage
	Const pgsize = 3	
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

	SQL = SQL & "DECLARE @CNT int															"
	SQL = SQL & "DECLARE @SNUM int															"
	SQL = SQL & "DECLARE @ENUM int															"
	SQL = SQL & "SET @ENUM = " & page & " * " & pgsize & "									"
	SQL = SQL & "SET @SNUM = (@ENUM - " & pgsize & ")+1										"

	SQL = SQL & "SELECT																		"
	SQL = SQL & "	@CNT = COUNT(*)															"
	SQL = SQL & "FROM																		"
	SQL = SQL & "	T_BOARDMAST AS BM														"
	SQL = SQL & "WHERE																		"
	SQL = SQL & "	1 = 1																	"
If sYN_USE <> "" Then
	SQL = SQL & "	AND YN_USE = '" & sYN_USE & "'											"
End If
If sKEYWORD <> "" Then
	SQL = SQL & "	AND NM_BOARDNM LIKE '%'+'" & sKEYWORD & "'+'%'							"
End If

	SQL = SQL & "SELECT																		"
	SQL = SQL & "	@CNT AS CNT 															"
	SQL = SQL & "	, CD_BOARDCD, NM_BOARDNM, NM_BOARDNM, YN_USE, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, NO_VIEWORDER																	"
	'					1				2			3		4		5			6			7		8			9
	SQL = SQL & "	, NM_FIELDNM_1, NM_FIELDNM_2, NM_FIELDNM_3, NM_FIELDNM_4, NM_FIELDNM_5, NM_FIELDNM_6, NM_FIELDNM_7, NM_FIELDNM_8, NM_FIELDNM_9, NM_FIELDNM_10						"
	'						10			11				12				13			14			15				16				17			18			19
	SQL = SQL & "	, NM_FIELDNM_11, NM_FIELDNM_12, NM_FIELDNM_13, NM_FIELDNM_14, NM_FIELDNM_15																							"
	'						20				21			22				23				24				
	SQL = SQL & "	, NM_FIELDINFO_1, NM_FIELDINFO_2, NM_FIELDINFO_3, NM_FIELDINFO_4, NM_FIELDINFO_5, NM_FIELDINFO_6, NM_FIELDINFO_7, NM_FIELDINFO_8, NM_FIELDINFO_9, NM_FIELDINFO_10	"
	'						25				26				27				28				29				30				31				32				33				34
	SQL = SQL & "	, NM_FIELDINFO_11, NM_FIELDINFO_12, NM_FIELDINFO_13, NM_FIELDINFO_14, NM_FIELDINFO_15																				"
	'						35				36				37					38				39
	SQL = SQL & "	, CD_FIELDTYPE_1, CD_FIELDTYPE_2, CD_FIELDTYPE_3, CD_FIELDTYPE_4, CD_FIELDTYPE_5, CD_FIELDTYPE_6, CD_FIELDTYPE_7, CD_FIELDTYPE_8 ,CD_FIELDTYPE_9, CD_FIELDTYPE_10	"
	'						40				41					42			43				44				45				46				47				48				49
	SQL = SQL & "	, CD_FIELDTYPE_11, CD_FIELDTYPE_12, CD_FIELDTYPE_13, CD_FIELDTYPE_14, CD_FIELDTYPE_15																				"
	'						50				51					52				53				54

	SQL = SQL & "	, CD_REQUIRED_1, CD_REQUIRED_2, CD_REQUIRED_3, CD_REQUIRED_4, CD_REQUIRED_5, CD_REQUIRED_6, CD_REQUIRED_7, CD_REQUIRED_8 ,CD_REQUIRED_9, CD_REQUIRED_10	"
	'						55				56					57			58				59				60				61				62				63				64
	SQL = SQL & "	, CD_REQUIRED_11, CD_REQUIRED_12, CD_REQUIRED_13, CD_REQUIRED_14, CD_REQUIRED_15																				"
	'						65				66					67				68				69

	SQL = SQL & "	, CD_INUSER, NO_INIPADDR, DT_INSYSDATE, CD_MDUSER, NO_MDIPADDR, DT_MDSYSDATE, CD_BOARDTYPE, YN_DOWNLOAD																"
	'						70			71			72			73			74			75				76			77
	SQL = SQL & "FROM																		"
	SQL = SQL & "	(																		"
	SQL = SQL & "	SELECT ROW_NUMBER() OVER (ORDER BY CD_BOARDCD DESC) AS [No], *			"
	SQL = SQL & "	FROM T_BOARDMAST 														"
	SQL = SQL & "	WHERE 1 = 1																"

If sYN_USE <> "" Then
	SQL = SQL & "	AND YN_USE = '" & sYN_USE & "'											"
End If

If sKEYWORD <> "" Then
	SQL = SQL & "	AND NM_BOARDNM LIKE '%'+'" & sKEYWORD & "'+'%'							"
End If

	SQL = SQL & ") J																		"
	SQL = SQL & "WHERE																		"
	SQL = SQL & "	No BETWEEN @SNUM AND @ENUM												"
	SQL = SQL & "ORDER BY																	"
	SQL = SQL & "	NO ASC																	"
	objADO.setSql(SQL)
	AryList = objADO.getArrRs()

	If IsArray(AryList) Then
	'If Not AryList.Eof then
		tn =  AryList(0,0)							'총 페이지수
		pgcount = - INT( - ( tn/pgsize ) )			'페이지수 계산
		If (page>pgcount) Then page=pgcount End If 	'페이지수 이상의 페이지를 요구하면 마지막 페이지로

		'이전page 다음page
		prevPage = page - ((page-1) Mod setsize) - setsize
		IF prevPage < 1 THEN
			prevPage = 0
		END IF
		nextPage = page - ((page-1) Mod setsize) + setsize
		IF nextPage > pgcount THEN
			nextPage = 0
		END If
	Else
		tn = 0
		pgcount = 0
	End If

SQL = ""
SQL = SQL & "SELECT CD_CODE, NM_CODE, NM_CODE1				"
SQL = SQL & "FROM T_CODE                                    "
SQL = SQL & "WHERE YN_USE = 'Y' AND CD_GUBUN = 'DATATYPE'   "
SQL = SQL & "ORDER BY NO_ORDER ASC                          "
objADO.setSql(SQL)
AryDataType = objADO.getArrRs()

SQL = ""
SQL = SQL & "SELECT CD_CODE, NM_CODE, NM_CODE1				"
SQL = SQL & "FROM T_CODE                                    "
SQL = SQL & "WHERE YN_USE = 'Y' AND CD_GUBUN = 'BOARDTYPE'  "
SQL = SQL & "ORDER BY NO_ORDER ASC                          "
objADO.setSql(SQL)
AryBoardType = objADO.getArrRs()

SQL = ""
SQL = SQL & "SELECT CD_CODE, NM_CODE, NM_CODE1				"
SQL = SQL & "FROM T_CODE                                    "
SQL = SQL & "WHERE YN_USE = 'Y' AND CD_GUBUN = 'IS_REQUIRED'   "
SQL = SQL & "ORDER BY NO_ORDER ASC                          "
objADO.setSql(SQL)
AryRequireType = objADO.getArrRs()

Set objADO = Nothing
%>
<!doctype html>
<html>
<head>
	<!--#include virtual="/manager/inc/inc.ui.head.asp"-->
</head>
<body>
	<div id="wrapper">
        <!--#include virtual="/manager/inc/inc.ui.navi.asp"-->
		
		<div id="page-wrapper" class="gray-bg">
			
			<!--#include virtual="/manager/inc/inc.ui.top.asp"-->

			<div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-10">
                    <h2>System Board (MASTER)</h2>
                    <ol class="breadcrumb">
                        <li>
                            <a href="/manager/">Home</a>
                        </li>
                        <li>
                            <a href="javascript:void(0)">시스템 관리</a>
                        </li>
						<li>
                            <a href="javascript:void(0)">게시판 관리</a>
                        </li>
                        <li class="active">
                            <strong>System Board (MASTER)</strong>
                        </li>
                    </ol>
                </div>
            </div>

            <div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-12">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-search"></i> 검색 <small>게시판 마스터 정보를 검색합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<div class="form-group">
									<label>사용여부</label>
									<select class="form-control " name="sYN_USE">
										<option value="">전체</option>
										<option value="Y" <%=getboolean(sYN_USE = "Y", "selected='selected'", "")%>>예</option>
										<option value="N" <%=getboolean(sYN_USE = "N", "selected='selected'", "")%>>아니오</option>
									</select>
								</div>
								<div class="form-group">
									<label>게시판명</label>
									<input type="text" class="form-control" id="sKEYWORD" name="sKEYWORD" value="<%=sKEYWORD%>" placeholder="검색하실 게시판명을 입력해주세요." />
									
								</div>
								<div class="form-group">
									<button type="button" class="btn btn-default btn-sm" onclick="fn_Search();"><i class="fa fa-search"></i> 조회</button>	
									<button type="button" class="btn btn-info btn-sm" onclick="fn_insview();"><i class="fa fa-plus"></i> 등록</button>	
								</div>
							</div>
						</div>

						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-list"></i> 게시판 리스트 <small>게시판 마스터 리스트입니다.</small></h5>
							</div>
							<div class="ibox-content">

								<div class="table-responsive">
									<form class="form-group-sm" id="frm_BoardMast" name="frm_BoardMast">											
										<input type="hidden" id="hid_mode" name="hid_mode" />
										<input type="hidden" id="CD_BOARDCD" name="CD_BOARDCD" />
										
										<div id="ins_board" style="display:none;">

											<table class="table table-hover table-striped">
											<thead>
												<tr>
													<th class="text-center">게시판명</th>
													<th class="text-center">타입</th>
													<th class="text-center">사용여부</th>
													<th class="text-center">조회수</th>
													<th class="text-center">답글</th>
													<th class="text-center">덧글</th>
													<th class="text-center">다운로드</th>
													<th class="text-center">XML배포</th>
													<th class="text-center">정렬순서</th>
													<th class="text-center">저장</th>
												</tr>
											</thead>
											<tbody>
												<tr>
													<td>
														<input class="form-control" type="text" id="NM_BOARDNM_0" name="NM_BOARDNM_0" size="30">
													</td>
													<td>
														<select class="form-control" id="CD_BOARDTYPE_0" name="CD_BOARDTYPE_0">
														<%
														if isarray(AryBoardType) then 
															for jloop=0 to ubound(AryBoardType,2)
														%>
															<option value="<%=AryBoardType(0,jloop)%>"><%=AryBoardType(1,jloop)%></option>
														<%
															next
														end if 
														%>
														</select>
													</td>
													<td>
														<select class="form-control" id="YN_USE_0" name="YN_USE_0">
															<option value="Y" selected>사용</option>
															<option value="N">사용안함</option>
														</select>
													</td>
													<td>
														<select class="form-control" id="YN_CNTVIEW_0" name="YN_CNTVIEW_0">
															<option value="Y" selected>사용</option>
															<option value="N">사용안함</option>
														</select>
													</td>
													<td>
														<select class="form-control" id="YN_ANSWER_0" name="YN_ANSWER_0">
															<option value="Y">사용</option>
															<option value="N" selected>사용안함</option>
														</select>
													</td>
													<td>
														<select class="form-control" id="YN_REPLY_0" name="YN_REPLY_0">
															<option value="Y">사용</option>
															<option value="N" selected>사용안함</option>
														</select>
													</td>
													<td>
														<select class="form-control" id="YN_DOWNLOAD_0" name="YN_DOWNLOAD_0">
															<option value="Y">사용</option>
															<option value="N" selected>사용안함</option>
														</select>
													</td>
													<td>
														<select class="form-control" id="YN_XML_0" name="YN_XML_0">
															<option value="Y">사용</option>
															<option value="N" selected>사용안함</option>
														</select>
													</td>
													<td>
														<input class="form-control" type="text" id="NO_VIEWORDER_0" name="NO_VIEWORDER_0" value="0" size="5">
													</td>
													<td>
														<button type="button" class="btn btn-info btn-sm" onclick="fn_Edit('I','0');"><i class="fa fa-plus"></i> 추가</button>	
													</td>
												</tr>

												<%for iloop=1 to 15%>
												<tr>
													<td colspan="10">
														<table class="table">
														<tr>
															<td>
																<input type="checkbox" id="ck_field_<%=iloop%>" name="ck_field_<%=iloop%>" onclick="fn_field_click('<%=iloop%>')">필드<%=iloop%>
															</td>
															<td>
																<div id="field_<%=iloop%>" style="display:none;">
																	&nbsp;&nbsp;&nbsp;필드명&nbsp;<input class="form-control" type="text" id="NM_FIELDNM<%=iloop%>_0" name="NM_FIELDNM<%=iloop%>_0" size="20">
																	&nbsp;&nbsp;&nbsp;설명&nbsp;<input class="form-control" type="text" id="NM_FIELDINFO<%=iloop%>_0" name="NM_FIELDINFO<%=iloop%>_0" size="40">
																	&nbsp;&nbsp;&nbsp;타입&nbsp;
																	<select class="form-control" id="CD_FIELDTYPE<%=iloop%>_0" name="CD_FIELDTYPE<%=iloop%>_0">
																		<%
																		if isarray(AryDataType) then 
																		%>
																		<option value=""></option>
																		<%
																			for jloop=0 to ubound(AryDataType,2)
																		%>
																		<option value="<%=AryDataType(0,jloop)%>"><%=AryDataType(1,jloop)%></option>
																		<%
																			next 
																		end if 
																		%>
																	</select>
																	&nbsp;&nbsp;&nbsp;필수여부&nbsp;
																	<select class="form-control" id="CD_REQUIRED<%=iloop%>_0" name="CD_REQUIRED<%=iloop%>_0">
																		<%
																		if isarray(AryRequireType) then 
																		%>
																		<option value=""></option>
																		<%
																			for jloop=0 to ubound(AryRequireType,2)
																		%>
																		<option value="<%=AryRequireType(0,jloop)%>"><%=AryRequireType(1,jloop)%></option>
																		<%
																			next 
																		end if 
																		%>
																	</select>
																</div>
															</td>
														</tr>
														</table>
													</td>
												</tr>
												<%next%>
											</tbody>
											</table>
										</div>



										<div id="list_board" style="display:;">

										<table class="table table-striped table-hover">
										<thead>
											<tr>
												<th class="text-center">게시판ID</th>
												<th class="text-center">게시판명</th>
												<th class="text-center">타입</th>
												<th class="text-center">사용여부</th>
												<th class="text-center">조회수</th>
												<th class="text-center">답글</th>
												<th class="text-center">덧글</th>
												<th class="text-center">다운로드</th>
												<th class="text-center">XML배포</th>
												<th class="text-center">정렬순서</th>
												<th class="text-center">저장</th>
											</tr>
										</thead>
									<%
										If isarray(AryList) Then 
											For iloop=0 To Ubound(AryList,2)
									%>
										<tbody>
											<tr>
												<td>
													<a href="javascript:GoBoard(<%=AryList(1,iloop)%>)"><%=AryList(1,iloop)%></a>
												</td>
												<td><input class="form-control" type="text" id="NM_BOARDNM_<%=AryList(1,iloop)%>" name="NM_BOARDNM_<%=AryList(1,iloop)%>" value="<%=getHtml(AryList(2,iloop))%>" size="30"></td>
												<td>
													<select class="form-control" id="CD_BOARDTYPE_<%=AryList(1,iloop)%>" name="CD_BOARDTYPE_<%=AryList(1,iloop)%>">
												<%
												if isarray(AryBoardType) then 
													for jloop=0 to ubound(AryBoardType,2)
												%>
														<option value="<%=AryBoardType(0,jloop)%>" <%=getBoolean(AryList(76,iloop)=AryBoardType(0,jloop),"selected","")%>><%=AryBoardType(1,jloop)%></option>
												<%
													next
												end if 
												%>
													</select>
												</td>
												<td>
													<select class="form-control" id="YN_USE_<%=AryList(1,iloop)%>" name="YN_USE_<%=AryList(1,iloop)%>">
														<option value="Y" <%=getBoolean(AryList(4,iloop)="Y","selected","")%>>사용</option>
														<option value="N" <%=getBoolean(AryList(4,iloop)="N","selected","")%>>사용안함</option>
													</select>
												</td>
												<td>
													<select class="form-control" id="YN_CNTVIEW_<%=AryList(1,iloop)%>" name="YN_CNTVIEW_<%=AryList(1,iloop)%>">
														<option value="Y" <%=getBoolean(AryList(5,iloop)="Y","selected","")%>>사용</option>
														<option value="N" <%=getBoolean(AryList(5,iloop)="N","selected","")%>>사용안함</option>
													</select>
												</td>
												<td>
													<select class="form-control" id="YN_ANSWER_<%=AryList(1,iloop)%>" name="YN_ANSWER_<%=AryList(1,iloop)%>">
														<option value="Y" <%=getBoolean(AryList(6,iloop)="Y","selected","")%>>사용</option>
														<option value="N" <%=getBoolean(AryList(6,iloop)="N","selected","")%>>사용안함</option>
													</select>
												</td>
												<td>
													<select class="form-control" id="YN_REPLY_<%=AryList(1,iloop)%>" name="YN_REPLY_<%=AryList(1,iloop)%>">
														<option value="Y" <%=getBoolean(AryList(7,iloop)="Y","selected","")%>>사용</option>
														<option value="N" <%=getBoolean(AryList(7,iloop)="N","selected","")%>>사용안함</option>
													</select>
												</td>
												<td>
													<select class="form-control" id="YN_DOWNLOAD_<%=AryList(1,iloop)%>" name="YN_DOWNLOAD_<%=AryList(1,iloop)%>">
														<option value="Y" <%=getBoolean(AryList(77,iloop)="Y","selected","")%>>사용</option>
														<option value="N" <%=getBoolean(AryList(77,iloop)="N","selected","")%>>사용안함</option>
													</select>
												</td>
												<td>
													<select class="form-control" id="YN_XML_<%=AryList(1,iloop)%>" name="YN_XML_<%=AryList(1,iloop)%>">
														<option value="Y" <%=getBoolean(AryList(8,iloop)="Y","selected","")%>>사용</option>
														<option value="N" <%=getBoolean(AryList(8,iloop)="N","selected","")%>>사용안함</option>
													</select>
												</td>
												<td>
													<input class="form-control" type="text" id="NO_VIEWORDER_<%=AryList(1,iloop)%>" name="NO_VIEWORDER_<%=AryList(1,iloop)%>" value="<%=AryList(9,iloop)%>" size="5">
												</td>
												<td rowspan="2">
													<button type="button" class="btn btn-info btn-sm" onclick="fn_Edit('E','<%=AryList(1,iloop)%>');"><i class="fa fa-save"></i> 저장</button>	
												</td>
											</tr>
											<tr>
												<td colspan="10">
													<table class="table">
														<tr>
														<%for jloop=1 to 8%>
															<td>
																[필드<%=jloop%>]
																<input class="form-control" type="text" id="NM_FIELDNM<%=jloop%>_<%=AryList(1,iloop)%>" name="NM_FIELDNM<%=jloop%>_<%=AryList(1,iloop)%>" value="<%=AryList(9+jloop,iloop)%>" size="17">
																[설명]
																<input class="form-control" type="text" id="NM_FIELDINFO<%=jloop%>_<%=AryList(1,iloop)%>" name="NM_FIELDINFO<%=jloop%>_<%=AryList(1,iloop)%>" value="<%=AryList(24+jloop,iloop)%>" size="17">
																[타입]
																<select class="form-control" id="CD_FIELDTYPE<%=jloop%>_<%=AryList(1,iloop)%>" name="CD_FIELDTYPE<%=jloop%>_<%=AryList(1,iloop)%>">
																<%
																if isarray(AryDataType) then 
																%>
																	<option value="" <%=getBoolean(AryList(39+jloop,iloop)="","selected","")%>></option>
																<%
																for kloop=0 to ubound(AryDataType,2)
																%>
																	<option value="<%=AryDataType(0,kloop)%>" <%=getBoolean(AryList(39+jloop,iloop)=AryDataType(0,kloop),"selected","")%>><%=AryDataType(1,kloop)%></option>
																<%
																next 
																end if 
																%>
																</select>
																[필수여부]
																<select class="form-control" id="CD_REQUIRED<%=jloop%>_<%=AryList(1,iloop)%>" name="CD_REQUIRED<%=jloop%>_<%=AryList(1,iloop)%>">
																<%
																if isarray(AryRequireType) then 
																%>
																	<option value="" <%=getBoolean(AryList(54+jloop,iloop)="","selected","")%>></option>
																<%
																for kloop=0 to ubound(AryRequireType,2)
																%>
																	<option value="<%=AryRequireType(0,kloop)%>" <%=getBoolean(AryList(54+jloop,iloop)=AryRequireType(0,kloop),"selected","")%>><%=AryRequireType(1,kloop)%></option>
																<%
																next 
																end if 
																%>
																</select>
															</td>
														<%next%>
														</tr>

														<tr>
														<%for jloop=9 to 15%>
															<td>
																[필드<%=jloop%>]
																<input class="form-control" type="text" id="NM_FIELDNM<%=jloop%>_<%=AryList(1,iloop)%>" name="NM_FIELDNM<%=jloop%>_<%=AryList(1,iloop)%>" value="<%=AryList(9+jloop,iloop)%>" size="17">
																[설명]
																<input class="form-control" type="text" id="NM_FIELDINFO<%=jloop%>_<%=AryList(1,iloop)%>" name="NM_FIELDINFO<%=jloop%>_<%=AryList(1,iloop)%>" value="<%=AryList(24+jloop,iloop)%>" size="17">
																[타입]
																<select class="form-control" id="CD_FIELDTYPE<%=jloop%>_<%=AryList(1,iloop)%>" name="CD_FIELDTYPE<%=jloop%>_<%=AryList(1,iloop)%>">
																<%
																if isarray(AryDataType) then 
																%>
																	<option value="" <%=getBoolean(AryList(39+jloop,iloop)="","selected","")%>></option>
																<%
																for kloop=0 to ubound(AryDataType,2)
																%>
																	<option value="<%=AryDataType(0,kloop)%>" <%=getBoolean(AryList(39+jloop,iloop)=AryDataType(0,kloop),"selected","")%>><%=AryDataType(1,kloop)%></option>
																<%
																next 
																end if 
																%>
																</select>	
																[필수여부]
																<select class="form-control" id="CD_REQUIRED<%=jloop%>_<%=AryList(1,iloop)%>" name="CD_REQUIRED<%=jloop%>_<%=AryList(1,iloop)%>">
																<%
																if isarray(AryRequireType) then 
																%>
																	<option value="" <%=getBoolean(AryList(54+jloop,iloop)="","selected","")%>></option>
																<%
																for kloop=0 to ubound(AryRequireType,2)
																%>
																	<option value="<%=AryRequireType(0,kloop)%>" <%=getBoolean(AryList(54+jloop,iloop)=AryRequireType(0,kloop),"selected","")%>><%=AryRequireType(1,kloop)%></option>
																<%
																next 
																end if 
																%>
																</select>
															</td>
														<%next%>
														</tr>
													</table>
												</td>
											</tr>
										<%	Next	%>
											<tr>
												<td colspan="11" class="text-center">
													<%
														Call PageNavigation("/manager/board/BoardMast.asp","sYN_USE="&sYN_USE&"&sKEYWORD="&sKEYWORD, tn, pgcount, startpage, page, setsize, prevpage, nextpage)
													%>
												</td>
											</tr>
										<%
											End If 
										%>
										</tbody>
										</table>
										</div>
									</form>

								</div>
								<!-- /.table-responsive -->
							</div>
						</div>

					</div>
				</div>
			</div>


		</div>
		<!-- /#page-wrapper -->
	</div>
	<!-- /#wrapper -->

	<script type= "text/javascript">
		function fn_field_click(no) {
			if ($("#ck_field_"+no).prop("checked") == true) {
				$("#field_"+no).show();
			} else {
				$("#field_"+no).hide();
			}
		}

		function fn_Edit(mode, boardcd) {

			$("#hid_mode").val( mode );
			$("#CD_BOARDCD").val( boardcd );
			
			if ($("#NM_BOARDNM_"+boardcd).val() == "") {
				alert("게시판명을 입력하세요.");
				$("#NM_BOARDNM_"+boardcd).focus();
				return;
			}

			if (mode=="I") {
				for(var i=1;i<16;i++) {

					if ($("#ck_field_"+i).prop("checked") == true) {
						if ($("#NM_FIELDNM"+i+"_"+boardcd).val() == "") {
							alert("필드명을 입력하세요.");
							$("#NM_FIELDNM"+i+"_"+boardcd).focus();
							return;
						}
						if ($("#CD_FIELDTYPE"+i+"_"+boardcd+" option:selected").val() == "") {
							alert("타입을 선택하세요.");
							$("#CD_FIELDTYPE"+i+"_"+boardcd).focus();
							return;
						}
						if ($("#CD_REQUIRED"+i+"_"+boardcd+" option:selected").val() == "") {
							alert("필수여부를 선택하세요.");
							$("#CD_REQUIRED"+i+"_"+boardcd).focus();
							return;
						}
					}
				}
			} else {
				for(var i=1;i<16;i++) {
					if ($("#NM_FIELDNM"+i+"_"+boardcd).val() != "" || $("#NM_FIELDINFO"+i+"_"+boardcd).val() != "" || $("#CD_FIELDTYPE"+i+"_"+boardcd+" option:selected").val() != "" ) {
						if ($("#NM_FIELDNM"+i+"_"+boardcd).val() == "") {
							alert("필드명을 입력하세요.");
							$("#NM_FIELDNM"+i+"_"+boardcd).focus();
							return;
						}
						if ($("#CD_FIELDTYPE"+i+"_"+boardcd+" option:selected").val() == "") {
							alert("타입을 선택하세요.");
							$("#CD_FIELDTYPE"+i+"_"+boardcd).focus();
							return;
						}
						if ($("#CD_REQUIRED"+i+"_"+boardcd+" option:selected").val() == "") {
							alert("필수여부를 선택하세요.");
							$("#CD_REQUIRED"+i+"_"+boardcd).focus();
							return;
						}
					}
				}
			}
			document.frm_BoardMast.method = "post";
			document.frm_BoardMast.action = "BoardMast_Proc.asp"
			document.frm_BoardMast.submit();
		}

		function fn_insview() {
			var insboard = $("#ins_board");
			var list_board = $("#list_board");

			if (insboard.css("display") == "") { 
				insboard.hide(); 
				list_board.show(); 
			} else  {
				insboard.show();
				list_board.hide();
			}
		}

		function fn_Search() {
			//document.frm_BoardMast.method = "post";
			//document.frm_BoardMast.action = "BoardMast.asp";
			//document.frm_BoardMast.submit()
			
			var sYN_USE = $("#sYN_USE option:selected").val();
			var sKEYWORD = $("#sKEYWORD").val();
			
			document.location.href = "BoardMast.asp?sYN_USE="+sYN_USE+"&sKEYWORD="+sKEYWORD;
		}

		function GoBoard(CD_BOARDCD) {
			document.location.href="BoardDetail.asp?CD_BOARDCD="+CD_BOARDCD;
		}
	</script>	
</body>
</html>