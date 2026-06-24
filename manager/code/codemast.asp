<!-- #include virtual = "/Include/Config.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/code/codemast.asp
	'#  Create		: 2015.07.02
	'#	Info		: codemast 등록
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, arrRs, SQL, iloop

	Set objADO = new clsADO
	SQL = SQL & "	SELECT CD_GUBUN, CD_CODE, NM_CODE, NM_CODE1, NM_CODE2, NM_CODE3, NM_CODE4, NM_CODE5, NM_CODE6, NM_CODE7, NO_ORDER, YN_USE "
	SQL = SQL & "	FROM T_CODE				"
	SQL = SQL & "	WHERE CD_GUBUN = '*' 	"
	SQL = SQL & "	ORDER BY NO_ORDER ASC 	"

	objADO.setSql(SQL)
	arrRs= objADO.getArrRs()
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
		
		<div id="page-wrapper" class="gray-bg">
			
			<!--#include virtual="/manager/inc/inc.ui.top.asp"-->
			
			<div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-10">
                    <h2>System Code Data (MASTER)</h2>
                    <ol class="breadcrumb">
                        <li>
                            <a href="/manager/">Home</a>
                        </li>
                        <li>
                            <a href="javascript:void(0)">시스템 관리</a>
                        </li>
						<li>
                            <a href="javascript:void(0)">코드 관리</a>
                        </li>
                        <li class="active">
                            <strong>System Code Data (MASTER)</strong>
                        </li>
                    </ol>
                </div>
            </div>
			
			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-12">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-list"></i> 코드 마스터 관리 <small>코드 마스터 정보를 관리합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<div class="table-responsive">

									<form class="form-group-sm" id="frm_Code" name="frm_Code" method="post" action="codeProcessController.asp">
										<input type="hidden" id="MODE" name="MODE">
										<input type="hidden" id="RTN_STR" name="RTN_STR" value="codemast.asp">
										<input type="hidden" id="CD_GUBUN" name="CD_GUBUN" value="*">
										<input type="hidden" id="CD_CODE" name="CD_CODE">
										<input type="hidden" id="nm_string" name="nm_string">

										<table class="table table-striped table-hover">
											<thead>
												<tr>
													<th class="text-center">코드값</th>
													<th class="text-center">코드명</th>
													<th class="text-center">코드명1</th>
													<th class="text-center">코드명2</th>
													<th class="text-center">코드명3</th>
													<th class="text-center">코드명4</th>
													<th class="text-center">코드명5</th>
													<th class="text-center">코드명6</th>
													<th class="text-center">코드명7</th>
													<th class="text-center">정렬순서</th>
													<th class="text-center">사용여부</th>
													<th class="text-center">저장</th>
												</tr>
											</thead>
											<tbody>
												<tr>
													<td><input type="text" class="form-control" id="CD_CODE_0" name="CD_CODE_0" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE_0" name="NM_CODE_0" size="20"></td>
													<td><input type="text" class="form-control" id="NM_CODE1_0" name="NM_CODE1_0" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE2_0" name="NM_CODE2_0" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE3_0" name="NM_CODE3_0" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE4_0" name="NM_CODE4_0" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE5_0" name="NM_CODE5_0" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE6_0" name="NM_CODE6_0" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE7_0" name="NM_CODE7_0" size="10"></td>
													<td><input type="text" class="form-control text-right" id="NO_ORDER_0" name="NO_ORDER_0" size="10" value="0"></td>
													<td>
														<select class="form-control" name="YN_USE_0">
															<option value="Y" selected="selected">예</option>
															<option value="N">아니오</option>
														</select>
													</td>
													<td>
														<button type="button" class="btn btn-success btn-sm" onclick="f_Edit('I','0');"><i class="fa fa-plus"></i> 추가</button>
													</td>
												</tr>

											<%If IsArray(arrRs) Then %>
												<%For iloop = 0 To UBound(arrRs,2) %>
												<tr>
													<td>
														<input type="hidden" id="CD_CODE_<%=arrRs(1,iloop)%>" name="CD_CODE_<%=arrRs(1,iloop)%>" value="<%=arrRs(1,iloop)%>">
														<button type="button" class="btn btn-link btn-sm" onclick="f_goDetail('<%=arrRs(1,iloop)%>');"><%=arrRs(1,iloop)%></button>
													</td>
													<td><input type="text" class="form-control" id="NM_CODE_<%=arrRs(1,iloop)%>" name="NM_CODE_<%=arrRs(1,iloop)%>" value="<%=arrRs(2,iloop)%>" size="20"></td>
													<td><input type="text" class="form-control" id="NM_CODE1_<%=arrRs(1,iloop)%>" name="NM_CODE1_<%=arrRs(1,iloop)%>" value="<%=arrRs(3,iloop)%>" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE2_<%=arrRs(1,iloop)%>" name="NM_CODE2_<%=arrRs(1,iloop)%>" value="<%=arrRs(4,iloop)%>" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE3_<%=arrRs(1,iloop)%>" name="NM_CODE3_<%=arrRs(1,iloop)%>" value="<%=arrRs(5,iloop)%>" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE4_<%=arrRs(1,iloop)%>" name="NM_CODE4_<%=arrRs(1,iloop)%>" value="<%=arrRs(6,iloop)%>" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE5_<%=arrRs(1,iloop)%>" name="NM_CODE5_<%=arrRs(1,iloop)%>" value="<%=arrRs(7,iloop)%>" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE6_<%=arrRs(1,iloop)%>" name="NM_CODE6_<%=arrRs(1,iloop)%>" value="<%=arrRs(8,iloop)%>" size="10"></td>
													<td><input type="text" class="form-control" id="NM_CODE7_<%=arrRs(1,iloop)%>" name="NM_CODE7_<%=arrRs(1,iloop)%>" value="<%=arrRs(9,iloop)%>" size="10"></td>
													<td><input type="text" class="form-control text-right" id="NO_ORDER_<%=arrRs(1,iloop)%>" name="NO_ORDER_<%=arrRs(1,iloop)%>" value="<%=arrRs(10,iloop)%>" size="10"></td>
													<td>
														<select class="form-control" name="YN_USE_<%=arrRs(1,iloop)%>">
															<option value="Y" <%=getBoolean(arrRs(11,iloop) = "Y", "selected='selected'", "")%>>예</option>
															<option value="N" <%=getBoolean(arrRs(11,iloop) = "N", "selected='selected'", "")%>>아니오</option>
														</select>
													</td>
													<td>
														<button type="button" class="btn btn-info btn-sm" onclick="f_Edit('U','<%=arrRs(1,iloop)%>');"><i class="fa fa-check"></i> 수정</button>
													</td>
												</tr>
												<%Next %>
											<%End If %>

											</tbody>
										</table>
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

	<script type="text/javascript">
		function f_Edit(MODE, CD_CODE) {
			var CD_CODE;
			var NM_CODE;
			var NO_ORDER;

			if (MODE=="I") {
				CD_CODE = $("#CD_CODE_0");
				NM_CODE = $("#NM_CODE_0");
				NO_ORDER = $("#NO_ORDER_0");

				if ($.trim(CD_CODE.val()) == "") {
					alert("코드정보를 입력하세요.");
					CD_CODE.focus();
					return;
				}

				if ($.trim(NM_CODE.val()) == "") {
					alert("코드명을 입력하세요.");
					NM_CODE.focus();
					return;
				}

				/*
				if(ValEng_V3(CD_CODE.val(), "", true) == false) { 
					alert("코드정보를 영어로만 입력해주세요.");
					return; 
				};
				*/
				if(ValNumber_V3(NO_ORDER.val(), "", true) == false) { 
					alert("정렬순서를 숫자로만 입력해주세요.");
					return; 
				};
				
				$("#CD_CODE").val( $("#CD_CODE_0").val() );
				$("#nm_string").val( "0" );

			} else if (MODE=="U") {

				NM_CODE = $("#NM_CODE_" + CD_CODE);
				NO_ORDER = $("#NO_ORDER_" + CD_CODE);

				if ($.trim(NM_CODE.val()) == "") {
					alert("코드명을 입력하세요.");
					NM_CODE.focus();
					return;
				}

				if(ValNumber_V3(NO_ORDER.val(), "정렬순서를", true) == false) { return false; };

				$("#CD_CODE").val( CD_CODE );
				$("#nm_string").val( CD_CODE );
			}

			$("#MODE").val( MODE );
			document.frm_Code.submit();
		}

		function f_goDetail(sGUBUN) {
			document.location.href="codedetail.asp?sGUBUN="+sGUBUN;
		}
	</script>
</body>
</html>