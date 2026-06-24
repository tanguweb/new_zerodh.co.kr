<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_XMLpath.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/board/BoardDetailView.asp
	'#  Create		: / 2010.09.13
	'#	Info		: 게시글보기 및 수정
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	
	Dim CD_BOARDID : CD_BOARDID = ReqQ("CD_BOARDID")

	Dim CD_BOARDCD, CD_BOARDKEY, NO_BOARD_DEPTH, NO_BOARD_SEQ
	Dim aryCodeGroupInfo, aryCodeInfo, aryBoardInfo, iloop, jloop, kloop
	Dim objADO, SQL
	Dim NM_BOARDNM, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, YN_DOWNLOAD
	Dim YN_DISPLAY,  NO_VIEWCNT, NO_VIEWORDER, NO_DOWNLOAD, YN_USE, NM_TITLE, NM_CONTENTS, CD_INUSER, DT_INSYSDATE
	Dim NM_XMLPATH, NM_XMLPATH2, NM_XMLQUERY, NM_XMLFILE
	Dim ary_FIELD(16), ary_FIELDNM(16), ary_FIELDINFO(16), ary_FIELDTYPE(16), ary_REQUIRED(16)
	Dim aryBeforeAfterImage

	Set objADO = new clsADO
	
	SQL = ""
	SQL = SQL & "SELECT																										"
	SQL = SQL & "	BM.NM_BOARDNM, BM.YN_USE, BM.YN_CNTVIEW, BM.YN_ANSWER, BM.YN_REPLY, BM.YN_XML							"	'5
	SQL = SQL & "	, BM.NM_FIELDNM_1, BM.NM_FIELDNM_2, BM.NM_FIELDNM_3, BM.NM_FIELDNM_4, BM.NM_FIELDNM_5					"	'10
	SQL = SQL & "	, BM.NM_FIELDNM_6, BM.NM_FIELDNM_7, BM.NM_FIELDNM_8, BM.NM_FIELDNM_9, BM.NM_FIELDNM_10					"	'15
	SQL = SQL & "	, BM.NM_FIELDNM_11, BM.NM_FIELDNM_12, BM.NM_FIELDNM_13, BM.NM_FIELDNM_14, BM.NM_FIELDNM_15				"	'20

	SQL = SQL & "	, BM.NM_FIELDINFO_1, BM.NM_FIELDINFO_2, BM.NM_FIELDINFO_3, BM.NM_FIELDINFO_4, BM.NM_FIELDINFO_5			"	'25
	SQL = SQL & "	, BM.NM_FIELDINFO_6, BM.NM_FIELDINFO_7, BM.NM_FIELDINFO_8, BM.NM_FIELDINFO_9, BM.NM_FIELDINFO_10		"	'30
	SQL = SQL & "	, BM.NM_FIELDINFO_11, BM.NM_FIELDINFO_12, BM.NM_FIELDINFO_13, BM.NM_FIELDINFO_14, BM.NM_FIELDINFO_15	"	'35

	SQL = SQL & "	, BM.CD_FIELDTYPE_1, BM.CD_FIELDTYPE_2, BM.CD_FIELDTYPE_3, BM.CD_FIELDTYPE_4, BM.CD_FIELDTYPE_5			"	'40
	SQL = SQL & "	, BM.CD_FIELDTYPE_6, BM.CD_FIELDTYPE_7, BM.CD_FIELDTYPE_8, BM.CD_FIELDTYPE_9, BM.CD_FIELDTYPE_10		"	'45
	SQL = SQL & "	, BM.CD_FIELDTYPE_11, BM.CD_FIELDTYPE_12, BM.CD_FIELDTYPE_13, BM.CD_FIELDTYPE_14, BM.CD_FIELDTYPE_15	"	'50
	
	SQL = SQL & "	, BM.CD_REQUIRED_1, BM.CD_REQUIRED_2, BM.CD_REQUIRED_3, BM.CD_REQUIRED_4, BM.CD_REQUIRED_5				"	'55
	SQL = SQL & "	, BM.CD_REQUIRED_6, BM.CD_REQUIRED_7, BM.CD_REQUIRED_8, BM.CD_REQUIRED_9, BM.CD_REQUIRED_10				"	'60
	SQL = SQL & "	, BM.CD_REQUIRED_11, BM.CD_REQUIRED_12, BM.CD_REQUIRED_13, BM.CD_REQUIRED_14, BM.CD_REQUIRED_15			"	'65

	SQL = SQL & "	, BD.CD_BOARDID, BD.NO_BOARD_DEPTH, BD.NO_BOARD_SEQ, BD.CD_BOARDCD, BD.YN_DISPLAY						"	'70
	SQL = SQL & "	, BD.NO_VIEWCNT, BD.NO_VIEWORDER, BD.YN_USE, BD.NM_TITLE, BD.NM_CONTENTS								"	'75
	
	SQL = SQL & "	, BD.NM_FIELD_1, BD.NM_FIELD_2, BD.NM_FIELD_3, BD.NM_FIELD_4, BD.NM_FIELD_5								"	'80
	SQL = SQL & "	, BD.NM_FIELD_6, BD.NM_FIELD_7, BD.NM_FIELD_8, BD.NM_FIELD_9, BD.NM_FIELD_10							"	'85
	SQL = SQL & "	, BD.NM_FIELD_11, BD.NM_FIELD_12, BD.NM_FIELD_13, BD.NM_FIELD_14, BD.NM_FIELD_15						"	'90

	SQL = SQL & "	, BD.CD_INUSER, BD.DT_INSYSDATE, BD.CD_BOARDKEY, BM.YN_DOWNLOAD, BD.NO_DOWNLOAD							"	'95
	SQL = SQL & "	, BM.NM_XMLPATH, BM.NM_XMLQUERY, BM.NM_XMLFILE															"	'98

	SQL = SQL & "FROM																										"
	SQL = SQL & "	T_BOARDDETAIL AS BD																						"
	SQL = SQL & "		INNER JOIN T_BOARDMAST AS BM ON BM.CD_BOARDCD = BD.CD_BOARDCD										"
	SQL = SQL & "WHERE																										"
	SQL = SQL & "	BD.YN_USE = 'Y' AND BD.CD_BOARDID = '" & CD_BOARDID & "' 												"

	objADO.setSql(SQL)
	aryBoardInfo = objADO.getArrRs()

	If not IsArray(aryBoardInfo) Then
		Response.Write "<script>alert('올바른 경로로 접근하세요.');history.back();</script>"
		Response.End
	End If
	
	CD_BOARDCD		= aryBoardInfo(69,0)
	CD_BOARDKEY		= aryBoardInfo(93,0)
	NO_BOARD_DEPTH	= aryBoardInfo(67,0)
	NO_BOARD_SEQ	= aryBoardInfo(68,0)

	NM_BOARDNM		= aryBoardInfo(0,0)
	YN_CNTVIEW		= aryBoardInfo(2,0)
	YN_ANSWER		= aryBoardInfo(3,0)
	YN_REPLY		= aryBoardInfo(4,0)
	YN_XML			= aryBoardInfo(5,0)
	YN_DOWNLOAD		= aryBoardInfo(94,0)

	YN_DISPLAY		= aryBoardInfo(70,0)
	NO_VIEWCNT		= aryBoardInfo(71,0)
	NO_VIEWORDER	= aryBoardInfo(72,0)
	NO_DOWNLOAD		= aryBoardInfo(95,0)
	YN_USE			= aryBoardInfo(73,0)
	NM_TITLE		= aryBoardInfo(74,0)
	NM_CONTENTS		= aryBoardInfo(75,0)
	CD_INUSER		= aryBoardInfo(91,0)
	DT_INSYSDATE	= aryBoardInfo(92,0)

	NM_XMLPATH		= aryBoardInfo(96,0)
	NM_XMLPATH2		= "/" & replace(aryBoardInfo(96,0),"\","/")
	NM_XMLQUERY		= aryBoardInfo(97,0)

for iloop = 1 to 15
	ary_FIELD(iloop)		= aryBoardInfo(iloop+75,0)
	ary_FIELDNM(iloop)		= aryBoardInfo(iloop+5,0)
	ary_FIELDINFO(iloop)	= aryBoardInfo(iloop+20,0)
	ary_FIELDTYPE(iloop)	= aryBoardInfo(iloop+35,0)
	ary_REQUIRED(iloop)		= aryBoardInfo(iloop+50,0)
next 


	SQL = "	SELECT CD_BOARDID, CD_BEFOREAFTER_ID, BEFORE_IMG, AFTER_IMG FROM T_BEFORE_AFTER_IMG WHERE CD_BOARDID = '" & CD_BOARDID & "' ORDER BY CD_BEFOREAFTER_ID ASC "
'	Response.write SQL
	objADO.setSql(SQL)
	aryBeforeAfterImage = objADO.getArrRs()
	


	Set objADO = Nothing
%>
<!doctype html>
<html>
<head>
	<!--#include virtual="/manager/inc/inc.ui.head.asp"-->
</head>
<body class="skin-2">	
	<div id="wrapper">

		<!--#include virtual="/manager/inc/inc.ui.navi.asp"-->

		<div id="page-wrapper" class="gray-bg dashbard-1">
			
			<!--#include virtual="/manager/inc/inc.ui.top.asp"-->
			


			<div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-10">
                    <h2><%=NM_BOARDNM%> 등록</h2>
                    <ol class="breadcrumb">
                        <li>
                            <a href="/manager/">Home</a>
                        </li>
                        <li>
                            <a href="javascript:void(0)">게시판 관리</a>
                        </li>
                        <li class="active">
                            <strong><%=NM_BOARDNM%> 등록</strong>
                        </li>
                    </ol>
                </div>
            </div>
			
			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-12">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-edit"></i> <%=NM_BOARDNM%> 등록 <small><%=NM_BOARDNM%> 게시물을 등록합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<form id="frm_boardins" name="frm_boardins" method="post" action="BeforeAfter_update_Proc.asp" ENCTYPE="multipart/form-data">
									<input type="hidden" id="MODE" name="MODE" value="U">
									<input type="hidden" id="CD_BOARDCD" name="CD_BOARDCD" value="<%=CD_BOARDCD%>">
									<input type="hidden" id="CD_BOARDID" name="CD_BOARDID" value="<%=CD_BOARDID%>">
									<input type="hidden" id="CD_BOARDKEY" name="CD_BOARDKEY" value="<%=CD_BOARDKEY%>">
									<input type="hidden" id="NO_BOARD_DEPTH" name="NO_BOARD_DEPTH" value="<%=NO_BOARD_DEPTH%>">
									<input type="hidden" id="NO_BOARD_SEQ" name="NO_BOARD_SEQ" value="<%=NO_BOARD_SEQ%>">
									<input type="hidden" id="YN_XML" name="YN_XML" value="<%=YN_XML%>">
									<input type="hidden" id="NM_XMLPATH" name="NM_XMLPATH" value="<%=NM_XMLPATH%>">
									<input type="hidden" id="NM_XMLFILE" name="NM_XMLFILE" value="<%=NM_XMLFILE%>">
									<input type="hidden" id="YN_USE" name="YN_USE" value="Y">
									<input type="hidden" id="arr_before_after_image" name="arr_before_after_image" value="">

									<table class="table table-hover table-striped">
									<tr>
										<td class="text-center"><strong>제목</strong></td>
										<td>
											<div class="col-lg-8">
												<input class="form-control" type="text" id="NM_TITLE" name="NM_TITLE" maxlength="200" value="<%=NM_TITLE%>" placeholder="제목을 입력해주세요." />
											</div>
										</td>
									</tr>
									<tr>
										<td class="text-center"><strong>내용</strong></td>
										<td>
											<div class="col-lg-12">
												<textarea id="NM_CONTENTS" name="NM_CONTENTS"><%=NM_CONTENTS%></textarea>
											</div>
										</td>
									</tr>
									<tr>
										<td class="text-center"><strong>전시여부</strong></td>
										<td>
											<div class="col-lg-3">
												<select class="form-control" id="YN_DISPLAY" name="YN_DISPLAY">
													<option value="Y" <%=getBoolean(YN_DISPLAY="Y","selected","")%>>예</option>
													<option value="N" <%=getBoolean(YN_DISPLAY="N","selected","")%>>아니오</option>
												</select>
											</div>
										</td>
									</tr>
								<%if NO_BOARD_DEPTH = 1 then %>
									<!--
									<tr>
										<td class="text-center"><strong>TOP 공지</strong></td>
										<td>
											<div class="col-lg-3">
												<select class="form-control" id="NO_VIEWORDER" name="NO_VIEWORDER">
													<option value="100" <%=getBoolean(NO_VIEWORDER="100","selected","")%>>아니오</option>
													<option value="0" <%=getBoolean(NO_VIEWORDER="0","selected","")%>>예</option>
												</select>
											</div>
											<div class="col-lg-9">
												<span class="text-info">*공지글은 언제나 최상단에 위치합니다.</span>
											</div>
										</td>
									</tr>-->
									<input type="hidden" id="NO_VIEWORDER" name="NO_VIEWORDER" value="100" />
								<%else %>
									<input type="hidden" id="NO_VIEWORDER" name="NO_VIEWORDER" value="<%=NO_VIEWORDER%>">
								<%end if %>

								<%if YN_CNTVIEW = "Y" then %>
									<tr>
										<td class="text-center"><strong>조회수</strong></td>
										<td>
											<div class="col-lg-3">
												<input class="form-control" type="text" id="NO_CNTVIEW" name="NO_CNTVIEW" value="<%=NO_VIEWCNT%>" maxlength="10" />
											</div>
										</td>
									</tr>
								<%else %>
									<input type="hidden" id="NO_CNTVIEW" name="NO_CNTVIEW" value="<%=NO_VIEWCNT%>">
								<%end if %>

								<%if YN_DOWNLOAD = "Y" then %>
									<tr>
										<td class="text-center"><strong>다운로드수</strong></td>
										<td>
											<div class="col-lg-3">
												<input class="form-control" type="text" id="NO_DOWNLOAD" name="NO_DOWNLOAD" value="<%=NO_DOWNLOAD%>" maxlength="5" size="5">
											</div>
										</td>
									</tr>
								<%else %>
									<input type="hidden" id="NO_DOWNLOAD" name="NO_DOWNLOAD" value="<%=NO_DOWNLOAD%>">
								<%end if %>

								<%
								for iloop=1 to 15
								if ary_FIELDNM(iloop) <> "" and ary_FIELDTYPE(iloop) <> "" then 
								%>
									<tr>
									<%if ary_FIELDTYPE(iloop) = "TEXT" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-4">
												<input class="form-control" type="text" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>" value="<%=ary_FIELD(iloop)%>" >
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-8">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%elseif ary_FIELDTYPE(iloop) = "BOOLEAN" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-4">
												<select class="form-control" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
													<option value="Y" <%=getBoolean(ary_FIELD(iloop)="Y","selected","")%>>예</option>
													<option value="N" <%=getBoolean(ary_FIELD(iloop)="N","selected","")%>>아니오</option>
												</select>
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-8">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>		
									<%elseif ary_FIELDTYPE(iloop) = "NUMBER" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-4">
												<input class="form-control" type="text" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>" value="<%=ary_FIELD(iloop)%>">
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-8">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%elseif ary_FIELDTYPE(iloop) = "FILE" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-4">
												<input type='file' id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
												&nbsp;&nbsp;<a href="<%=getboolean(YN_XML="N","/file/board/"&ary_FIELD(iloop),NM_XMLPATH2 & ary_FIELD(iloop))%>" target="_blank"><%=ary_FIELD(iloop)%></a>
												<%If Right(ary_FIELD(iloop),3) = "jpg" or Right(ary_FIELD(iloop),3) = "gif" Then%>
												<br><img src="<%=getboolean(YN_XML="N","/file/board/"&ary_FIELD(iloop),NM_XMLPATH2 & ary_FIELD(iloop))%>" border="0" height="100" align="left">
												<%End If%>
												<input type="hidden" id="hd_NM_FIELD_<%=iloop%>" name="hd_NM_FIELD_<%=iloop%>" value="<%=ary_FIELD(iloop)%>">
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-8">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%elseif ary_FIELDTYPE(iloop) = "CODE" then %>
										<%
											aryCodeGroupInfo = ""
											aryCodeGroupInfo = fn_CodeInfoWhere("*", " AND CD_CODE='"&ary_FIELDNM(iloop)&"'")
										%>
										<td class="text-center"><strong><%=aryCodeGroupInfo(2, 0)%></strong></td>
										<td>
											<div class="col-lg-4">
												<select class="form-control" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
										<%
											aryCodeInfo = ""
											aryCodeInfo = fn_CodeInfo(ary_FIELDNM(iloop))
											if isarray(aryCodeInfo) then 
											for jloop=0 to ubound(aryCodeInfo,2)
										%>				
													<option value="<%=aryCodeInfo(1,jloop)%>" <%=getBoolean(ary_FIELD(iloop)=aryCodeInfo(1,jloop),"selected","")%>><%=aryCodeInfo(2,jloop)%></option>
										<%
											next
											end if 
										%>
												</select>
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-8">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%else %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-4">
												<input class="form-control" type="text" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-8">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%end if %>
									</tr>
								<%
								end if 
								next 
								%>
									
									<tr>
										<td class="text-center"><strong>전후사진</strong></td>
										<td>
											<div class="col-lg-12">
											<table class="table table-bordered table-hover">
												<thead>
													<tr>
														<th class="text-center">
															<button type="button" class="btn btn-xs btn-success m-b-none BTN_BEFORE_AFTER_IMG_ADD" title="전후사진 업로드란을 추가합니다."><i class="fa fa-plus"></i></button>
														</th>
														<th class="text-center"><small>BEFORE</small></th>
														<th class="text-center"><small>AFTER</small></th>
													</tr>
												</thead>
												<tbody id="before_after_image">
												<%
												If IsArray(aryBeforeAfterImage) Then 
													For iloop = 0 To UBound(aryBeforeAfterImage,2)
												%>
													<tr>
														<td class="text-center">
															<button type="button" class="btn btn-xs btn-danger m-b-none BTN_BEFORE_AFTER_IMG_REMOVE" title="전후사진을 삭제합니다."><i class="fa fa-minus"></i></button>
														</td>
														<td>
															<input type="file" class="before_after_image" data-index="<%=iloop%>" name="BEFORE_IMG_<%=iloop%>" value="" />
															<input type="hidden" name="hd_BEFORE_IMG_<%=iloop%>" value="<%=aryBeforeAfterImage(2,iloop)%>" />

															<img src="/file/board/<%=aryBeforeAfterImage(2,iloop)%>" style="max-width: 100px; max-height: 100px"/>
														</td>
														<td>
															<input type="file" name="AFTER_IMG_<%=iloop%>" value="" />
															<input type="hidden" name="hd_AFTER_IMG_<%=iloop%>" value="<%=aryBeforeAfterImage(3,iloop)%>" />

															<img src="/file/board/<%=aryBeforeAfterImage(3,iloop)%>" style="max-width: 100px; max-height: 100px"/>

														</td>
													</tr>
												<%
													Next
												End If
												%>
												</thead>
											</table>											
											</div>
										</td>
									</tr>



								<%if YN_REPLY = "Y" then %>
									<tr>
										<td class="text-center"><strong>덧글</strong></td>
										<td><iframe src="BoardReply.asp?CD_BOARDID=<%=CD_BOARDID%>" width="100%" height="200"></iframe></td>
									</tr>
								<%end if %>
									</table>








									<p class="text-center">
										<button type="button" class="btn btn-info" onclick="fn_save();"> <i class="fa fa-check"></i> 저장</button>
										<button type="button" class="btn btn-danger" onclick="fn_del();"> <i class="fa fa-trash-o"></i> 삭제</button>
										<!--<input type="button" value="복사" onclick="fn_copy();">-->
										<button type="button" class="btn" onclick="document.location.href='BeforeAfterList.asp?CD_BOARDCD=<%=CD_BOARDCD%>'"> <i class="fa fa-list"></i> 목록</button>
									</p>
								</form>
							</div>
						</div>
					</div>
				</div>
			</div>

			<!--#include virtual="/manager/inc/inc.ui.footer.asp"-->
		</div>
	</div>

	<!--#include virtual="/manager/inc/inc.ui.hidden.asp"-->

<script type="text/javascript" src="/ckeditor/ckeditor.js"></script>
<script type="text/javascript">

	var before_after_img_idx = <%=UBound(aryBeforeAfterImage,2)%>;

	$(document).ready(function() {
		CKEDITOR.replace('NM_CONTENTS',
		{
			toolbar: 'Full',
			skin: 'kama', //에디터 스킨 (office2013, kama, moonocolor, moono-dark)
			//uiColor: '#14B8C4',
			enterMode: '1', //엔터키 태그 1:<p>, 2:<br>, 3:<div>
			shiftEnterMode: '2', //쉬프트+엔터키 태그 1:<p>, 2:<br>, 3:<div>
			height:"400px",
			resize_enabled:true,
			filebrowserUploadUrl: '/function/fn_ckeditorFileUpload.asp?type=Files',
			filebrowserImageUploadUrl: '/function/fn_ckeditorFileUpload.asp?type=Images',
			filebrowserFlashUploadUrl: '/function/fn_ckeditorFileUpload.asp?type=Flash'
		});

		$('.BTN_BEFORE_AFTER_IMG_ADD').click(function() {
			
			before_after_img_idx++;

			var html = '';
			html += '<tr>';
			html += '	<td class="text-center">';
			html += '		<button type="button" class="btn btn-xs btn-danger m-b-none BTN_BEFORE_AFTER_IMG_REMOVE" title="전후사진을 삭제합니다."><i class="fa fa-minus"></i></button>';
			html += '	</td>';
			html += '	<td>';
			html += '		<input type="file" class="before_after_image" data-index="'+before_after_img_idx+'" name="BEFORE_IMG_'+before_after_img_idx+'" value="" />';
			html += '		<input type="hidden" name="hd_BEFORE_IMG_'+before_after_img_idx+'" value="" />';
			html += '	</td>';
			html += '	<td>';
			html += '		<input type="file" name="AFTER_IMG_'+before_after_img_idx+'" value="" />';
			html += '		<input type="hidden" name="hd_AFTER_IMG_'+before_after_img_idx+'" value="" />';
			html += '	</td>';
			html += '</tr>';

			$('#before_after_image').append(html);
		});

	});

	
	$(document).on('click', '.BTN_BEFORE_AFTER_IMG_REMOVE', function() {
		if ($('#before_after_image').find('tr').length == 1) {
			alert('전후사진은 최소 1개 이상 업로드해주세요.');
		} else {
			$(this).parent('td').parent('tr').remove();
		}
	});


	function fn_save()
	{
		var NM_TITLE = document.getElementById("NM_TITLE");
		if (NM_TITLE.value == "") { alert("제목을 입력하세요"); NM_TITLE.focus(); return false; }

<%if YN_CNTVIEW = "Y" then %>
		var NO_CNTVIEW = document.getElementById("NO_CNTVIEW");
		if(ValNumber(NO_CNTVIEW, "정렬순서를", true) == false) { return false; };
<%end if %>

<%if YN_DOWNLOAD = "Y" then %>
		var NO_DOWNLOAD = document.getElementById("NO_DOWNLOAD");
		if(ValNumber(NO_DOWNLOAD, "다운로드수를", true) == false) { return false; };
<%end if %>

<%
for iloop=1 to 15
	if ary_FIELDNM(iloop) <> "" and ary_FIELDTYPE(iloop) <> "" and ary_FIELDTYPE(iloop) = "REQUIRED" then 
		if ary_FIELDTYPE(iloop) = "FILE" then 
%>
			if (document.getElementById("NM_FIELD_<%=iloop%>").value == "" && document.getElementById("hd_NM_FIELD_<%=iloop%>").value == "")
			{
				alert("<%=ary_FIELDNM(iloop)%>를 등록해 주세요.");
				document.getElementById("NM_FIELD_<%=iloop%>").focus();
				return false;
			}			
<%
		else
%>
			if (document.getElementById("NM_FIELD_<%=iloop%>").value == "")
			{
				alert("<%=ary_FIELDNM(iloop)%>를 등록해 주세요.");
				document.getElementById("NM_FIELD_<%=iloop%>").focus();
				return false;
			}			
<%
		end if 
	end if 
next 
%>
		
		var is_before_after_image = false;
		var arr_before_after_image = '';
		$('.before_after_image').each(function(i) {

			var idx = $(this).data('index');

			if (i == 0) {
				arr_before_after_image += idx;
			} else {
				arr_before_after_image += ',' + idx;
			}


			if ($('input[name="BEFORE_IMG_'+idx+'"]').val() == '' && $('input[name="hd_BEFORE_IMG_'+idx+'"]').val() == '') {
				alert('전후사진(BEFORE)을 업로드 해주세요.');
				$('input[name="BEFORE_IMG_'+idx+'"]').focus();
				is_before_after_image = false;
				return false;
			}

			if ($('input[name="AFTER_IMG_'+idx+'"]').val() == '' && $('input[name="hd_AFTER_IMG_'+idx+'"]').val() == '') {
				alert('전후사진(AFTER)을 업로드 해주세요.');
				$('input[name="AFTER_IMG_'+idx+'"]').focus();
				is_before_after_image = false;
				return false;
			}

			is_before_after_image = true;
		});
		
		//alert('확인중: [' + arr_before_after_image + ']');
//		return;

		if (is_before_after_image) {
			$('#arr_before_after_image').val(arr_before_after_image);
			$('#frm_boardins').attr('action', 'BeforeAfter_update_Proc.asp');
			document.frm_boardins.submit();
		}
	}

	function fn_del()
	{
		if(confirm("삭제 하시겠습니까?") == true)
		{
			$('#frm_boardins').attr('action', 'BeforeAfter_delete_Proc.asp');
			document.frm_boardins.submit();
		}
		else
		{
			return false;
		}
	}
</script>
</body>
</html>