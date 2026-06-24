<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/board/BoardDetailIns.asp
	'#  Create		: / 2010.09.08
	'#	Info		: 게시글등록
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim CD_BOARDCD : CD_BOARDCD = ReqQ("CD_BOARDCD")
	Dim CD_BOARDID : CD_BOARDID = ReqQ("CD_BOARDID")
	Dim CD_BOARDKEY : CD_BOARDKEY = ReqQ("CD_BOARDKEY")
	Dim NO_BOARD_DEPTH : NO_BOARD_DEPTH = ReqQ("NO_BOARD_DEPTH")
	Dim NO_BOARD_SEQ : NO_BOARD_SEQ = ReqQ("NO_BOARD_SEQ")
	Dim aryCodeGroupInfo, aryCodeInfo, aryBoardInfo, iloop, jloop, kloop
	Dim objADO, SQL
	Dim NM_BOARDNM, YN_USE, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML, YN_DOWNLOAD, NM_XMLPATH, NM_XMLQUERY, NM_XMLFILE
	Dim ary_FIELDNM(16), ary_FIELDINFO(16), ary_FIELDTYPE(16), ary_REQUIRED(16)

	Set objADO = new clsADO
	
	SQL = ""
	SQL = SQL & "SELECT																						"
	SQL = SQL & "	NM_BOARDNM, YN_USE, YN_CNTVIEW, YN_ANSWER, YN_REPLY, YN_XML								"	'5
	SQL = SQL & "	, NM_FIELDNM_1, NM_FIELDNM_2, NM_FIELDNM_3, NM_FIELDNM_4, NM_FIELDNM_5					"	'10
	SQL = SQL & "	, NM_FIELDNM_6, NM_FIELDNM_7, NM_FIELDNM_8, NM_FIELDNM_9, NM_FIELDNM_10					"	'15
	SQL = SQL & "	, NM_FIELDNM_11, NM_FIELDNM_12, NM_FIELDNM_13, NM_FIELDNM_14, NM_FIELDNM_15				"	'20
	SQL = SQL & "	, NM_FIELDINFO_1, NM_FIELDINFO_2, NM_FIELDINFO_3, NM_FIELDINFO_4, NM_FIELDINFO_5		"	'25
	SQL = SQL & "	, NM_FIELDINFO_6, NM_FIELDINFO_7, NM_FIELDINFO_8, NM_FIELDINFO_9, NM_FIELDINFO_10		"	'30
	SQL = SQL & "	, NM_FIELDINFO_11, NM_FIELDINFO_12, NM_FIELDINFO_13, NM_FIELDINFO_14, NM_FIELDINFO_15	"	'35
	SQL = SQL & "	, CD_FIELDTYPE_1, CD_FIELDTYPE_2, CD_FIELDTYPE_3, CD_FIELDTYPE_4, CD_FIELDTYPE_5		"	'40
	SQL = SQL & "	, CD_FIELDTYPE_6, CD_FIELDTYPE_7, CD_FIELDTYPE_8, CD_FIELDTYPE_9, CD_FIELDTYPE_10		"	'45
	SQL = SQL & "	, CD_FIELDTYPE_11, CD_FIELDTYPE_12, CD_FIELDTYPE_13, CD_FIELDTYPE_14, CD_FIELDTYPE_15	"	'50
	SQL = SQL & "	, CD_REQUIRED_1, CD_REQUIRED_2, CD_REQUIRED_3, CD_REQUIRED_4, CD_REQUIRED_5				"	'55
	SQL = SQL & "	, CD_REQUIRED_6, CD_REQUIRED_7, CD_REQUIRED_8, CD_REQUIRED_9, CD_REQUIRED_10			"	'60
	SQL = SQL & "	, CD_REQUIRED_11, CD_REQUIRED_12, CD_REQUIRED_13, CD_REQUIRED_14, CD_REQUIRED_15		"	'65
	SQL = SQL & "	, YN_DOWNLOAD,NM_XMLPATH,NM_XMLQUERY,NM_XMLFILE											"	'69
	SQL = SQL & "FROM																						"
	SQL = SQL & "	T_BOARDMAST																				"
	SQL = SQL & "WHERE																						"
	SQL = SQL & "	CD_BOARDCD = '" & CD_BOARDCD & "'														"
	objADO.setSql(SQL)
	aryBoardInfo = objADO.getArrRs()

	If not IsArray(aryBoardInfo) Then
		Response.Write "<script>alert('올바른 경로로 접근하세요.');history.back();</script>"
		Response.End
	End If

	NM_BOARDNM		= aryBoardInfo(0,0)
	YN_USE			= aryBoardInfo(1,0)
	YN_CNTVIEW		= aryBoardInfo(2,0)
	YN_ANSWER		= aryBoardInfo(3,0)
	YN_REPLY		= aryBoardInfo(4,0)
	YN_XML			= aryBoardInfo(5,0)
	NM_XMLPATH		= aryBoardInfo(37,0)
	NM_XMLQUERY		= aryBoardInfo(38,0)
	NM_XMLFILE		= aryBoardInfo(39,0)
	YN_DOWNLOAD		= aryBoardInfo(36,0)

	If YN_XML="Y" and (NM_XMLPATH="" or NM_XMLQUERY="" or NM_XMLFILE="") then
		Response.Write "<script>alert('XML설정후 등록가능합니다.');history.back();</script>"
		Response.End
	end if 

for iloop = 1 to 15
	ary_FIELDNM(iloop)		= aryBoardInfo(iloop+5,0)
	ary_FIELDINFO(iloop)	= aryBoardInfo(iloop+20,0)
	ary_FIELDTYPE(iloop)	= aryBoardInfo(iloop+35,0)
	ary_REQUIRED(iloop)		= aryBoardInfo(iloop+50,0)
next
	'########################################################################################
	'# CD_BOARDCD = 1016 게시판 전용 폼 커스터마이징 (내용/TOP공지/조회수 제거 + 입력칸 확대)
	'########################################################################################
	Dim bCustomForm, sColTitle, sColField, sColInfo
	bCustomForm = (CD_BOARDCD = "1012")
	If bCustomForm Then
		sColTitle = "12"
		sColField = "9"
		sColInfo  = "3"
	Else
		sColTitle = "8"
		sColField = "4"
		sColInfo  = "8"
	End If

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
						<!--
                        <li>
                            <a href="javascript:void(0)">게시판 관리</a>
                        </li>
						-->
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
								<form id="frm_boardins" name="frm_boardins" method="post" action="BoardDetail_Proc.asp" ENCTYPE="multipart/form-data">
									<input type="hidden" id="MODE" name="MODE" value="I">
									<input type="hidden" id="CD_BOARDCD" name="CD_BOARDCD" value="<%=CD_BOARDCD%>">
									<input type="hidden" id="CD_BOARDID" name="CD_BOARDID" value="<%=CD_BOARDID%>">
									<input type="hidden" id="NO_BOARDKEY" name="CD_BOARDKEY" value="<%=CD_BOARDKEY%>">
									<input type="hidden" id="NO_BOARD_DEPTH" name="NO_BOARD_DEPTH" value="<%=NO_BOARD_DEPTH%>">
									<input type="hidden" id="NO_BOARD_SEQ" name="NO_BOARD_SEQ" value="<%=NO_BOARD_SEQ%>">
									<input type="hidden" id="YN_XML" name="YN_XML" value="<%=YN_XML%>">
									<input type="hidden" id="NM_XMLPATH" name="NM_XMLPATH" value="<%=NM_XMLPATH%>">
									<input type="hidden" id="NM_XMLFILE" name="NM_XMLFILE" value="<%=NM_XMLFILE%>">
									<input type="hidden" id="YN_USE" name="YN_USE" value="Y">
									
									<table class="table table-hover table-striped">
									<tr>
										<td class="text-center"><strong>제목</strong></td>
										<td>
											<div class="col-lg-<%=sColTitle%>">
												<input class="form-control" type="text" id="NM_TITLE" name="NM_TITLE" maxlength="200" placeholder="제목을 입력해주세요." />
											</div>
										</td>
									</tr>
							<%if bCustomForm then %>
									<input type="hidden" id="NM_CONTENTS" name="NM_CONTENTS" value="">
								<%else %>                  
									<tr>
										<td class="text-center"><strong>내용</strong></td>
										<td>
											<div class="col-lg-12">
												<textarea id="NM_CONTENTS" name="NM_CONTENTS"></textarea>
											</div>
										</td>
									</tr>
								<%end if %>
									<tr>
										<td class="text-center"><strong>전시여부</strong></td>
										<td>
											<div class="col-lg-3">
												<select class="form-control" id="YN_DISPLAY" name="YN_DISPLAY">
													<option value="Y">예</option>
													<option value="N">아니오</option>
												</select>
											</div>
										</td>
									</tr>
							<%if bCustomForm then %>
									<input type="hidden" id="NO_VIEWORDER" name="NO_VIEWORDER" value="100">
									<input type="hidden" id="NO_CNTVIEW" name="NO_CNTVIEW" value="0">
								<%else %>                  
									<tr>
										<td class="text-center"><strong>TOP 공지</strong></td>
										<td>
											<div class="col-lg-3">
												<select class="form-control" id="NO_VIEWORDER" name="NO_VIEWORDER">
													<option value="100" >아니오</option>
													<option value="0" >예</option>
												</select>
											</div>
											<div class="col-lg-9">
												<p class="form-control-static text-info">*공지글은 언제나 최상단에 위치합니다.</p>
											</div>
										</td>
									</tr>

									<%if YN_CNTVIEW = "Y" then %>
									<tr>
										<td class="text-center"><strong>조회수</strong></td>
										<td>
											<div class="col-lg-3">
												<input class="form-control" type="text" id="NO_CNTVIEW" name="NO_CNTVIEW" value="0" maxlength="10" />
											</div>
										</td>
									</tr>
								  <%else %>
									<input type="hidden" id="NO_CNTVIEW" name="NO_CNTVIEW" value="0">
									<%end if %>
                  
								<%end if %>

								<%if YN_DOWNLOAD = "Y" then %>
									<tr>
										<td class="text-center"><strong>다운로드수</strong></td>
										<td>
											<div class="col-lg-3">
												<input class="form-control" type="text" id="NO_DOWNLOAD" name="NO_DOWNLOAD" value="0" maxlength="5" size="5">
											</div>
										</td>
									</tr>
								<%else %>
									<input type="hidden" id="NO_DOWNLOAD" name="NO_DOWNLOAD" value="0">
								<%end if %>

								<%
								for iloop=1 to 15
								if ary_FIELDNM(iloop) <> "" and ary_FIELDTYPE(iloop) <> "" then 
								%>
									<tr>
									<%if ary_FIELDTYPE(iloop) = "TEXT" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-<%=sColField%>">
												<input class="form-control" type="text" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>" value="" />
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-<%=sColInfo%>">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%elseif ary_FIELDTYPE(iloop) = "BOOLEAN" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-<%=sColField%>">
												<select class="form-control" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
													<option value="Y">예</option>
													<option value="N">아니오</option>
												</select>
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-<%=sColInfo%>">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>		
									<%elseif ary_FIELDTYPE(iloop) = "NUMBER" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-<%=sColField%>">
												<input class="form-control" type="text" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-<%=sColInfo%>">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%elseif ary_FIELDTYPE(iloop) = "FILE" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-<%=sColField%>">
												<input type='file' id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-<%=sColInfo%>">
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
											<div class="col-lg-<%=sColField%>">
												<select class="form-control" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
										<%
											aryCodeInfo = ""
											aryCodeInfo = fn_CodeInfo(ary_FIELDNM(iloop))
											if isarray(aryCodeInfo) then
											for jloop=0 to ubound(aryCodeInfo,2)
										%>				
													<option value="<%=aryCodeInfo(1,jloop)%>"><%=aryCodeInfo(2,jloop)%></option>
										<%
											next
											end if 
										%>
												</select>
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-<%=sColInfo%>">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%elseif ary_FIELDTYPE(iloop) = "DATE" then %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-<%=sColField%>">
												<div class="input-group date">
													<input class="form-control field_type_date" type="text" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>" value="" data-mask="99999999" />
													<span class="input-group-addon btnDate"><i class="fa fa-calendar"></i></span>
													<span class="input-group-addon btnDateRemove"><i class="fa fa-refresh"></i></span>
												</div>
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-<%=sColInfo%>">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%else %>
										<td class="text-center"><strong><%=ary_FIELDNM(iloop)%></strong></td>
										<td>
											<div class="col-lg-<%=sColField%>">
												<input class="form-control" type="text" id="NM_FIELD_<%=iloop%>" name="NM_FIELD_<%=iloop%>">
												<input type="hidden" id="CD_FIELDTYPE_<%=iloop%>" name="CD_FIELDTYPE_<%=iloop%>" value="<%=ary_FIELDTYPE(iloop)%>">
											</div>
											<div class="col-lg-<%=sColInfo%>">
												<p class="form-control-static text-info">*<%=ary_FIELDINFO(iloop)%></p>
											</div>
										</td>
									<%end if %>
									</tr>
								<%
								end if 
								next 
								%>
									</table>

									<p class="text-center">
										<button type="button" class="btn btn-info" onclick="fn_save();"> <i class="fa fa-check"></i> 저장</button>
										<button type="button" class="btn" onclick="document.location.href='BoardDetail.asp?CD_BOARDCD=<%=CD_BOARDCD%>'"> <i class="fa fa-list"></i> 목록</button>
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
	$(document).ready(function() {

		$('.field_type_date').datepicker({
			todayBtn: "linked",
			keyboardNavigation: false,
			forceParse: false,
			autoclose: true,
			format: 'yyyymmdd',
			keyboardNavigation: true,
			language: 'ko'
		});
		
		$('.btnDate').click(function() {
			$(this).prev('input.field_type_date').focus();
		});

		$('.btnDateRemove').click(function() {
			$(this).prev().prev('input.field_type_date').val("");
		});
	<%if not bCustomForm then %>

		CKEDITOR.replace('NM_CONTENTS',
		{
			toolbar: 'Full',
			skin: 'kama', //에디터 스킨 (office2013, kama, moonocolor, moono-dark)
			//uiColor: '#14B8C4',
			enterMode: '1', //엔터키 태그 1:<p>, 2:<br>, 3:<div>
			shiftEnterMode: '2', //쉬프트+엔터키 태그 1:<p>, 2:<br>, 3:<div>
			height:"400px",
			resize_enabled:true,
			allowedContent: true,
			filebrowserUploadUrl: '/function/fn_ckeditorFileUpload.asp?type=Files',
			filebrowserImageUploadUrl: '/function/fn_ckeditorFileUpload.asp?type=Images',
			filebrowserFlashUploadUrl: '/function/fn_ckeditorFileUpload.asp?type=Flash'
		});
    	<%end if %>

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
if ary_FIELDNM(iloop) <> "" and ary_FIELDTYPE(iloop) <> "" and ary_REQUIRED(iloop) = "REQUIRED" then 
%>
		if (document.getElementById("NM_FIELD_<%=iloop%>").value == "")
		{
			alert("<%=ary_FIELDNM(iloop)%>를 등록해 주세요.");
			document.getElementById("NM_FIELD_<%=iloop%>").focus();
			return false;
		}
<%
end if 
next 
%>
		document.frm_boardins.submit();
	}
</script>
</body>
</html>		