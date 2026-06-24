<!-- #include virtual = "/Include/Config.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/code/Codedetail.asp
	'#  Create		: 2010.09.08
	'#	Info		: codedetail 등록
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, Rs, ArrMast, ArrDetail, SQL, iloop, jloop
	Dim sGUBUN
	Dim CD_GUBUN, CD_CODE, NM_CODE, NM_CODE1, NM_CODE2, NM_CODE3, NM_CODE4, NM_CODE5, NM_CODE6, NM_CODE7, NO_ORDER, YN_USE

	sGUBUN = ReqQ("sGUBUN")

	Set objADO = new clsADO

	SQL = ""
	SQL = SQL & "	SELECT CD_CODE, NM_CODE					"
	SQL = SQL & "	FROM T_CODE								"
	SQL = SQL & "	WHERE CD_GUBUN = '*' AND YN_USE='Y' 	"
	SQL = SQL & "	ORDER BY NO_ORDER ASC 					"
	objADO.setSql(SQL)
	ArrMast = objADO.getArrRs()

	SQL = ""
	SQL = SQL & "	SELECT CD_GUBUN, CD_CODE, NM_CODE, NM_CODE1, NM_CODE2, NM_CODE3, NM_CODE4, NM_CODE5, NM_CODE6, NM_CODE7, NO_ORDER, YN_USE								"
	SQL = SQL & "	FROM T_CODE								"
	SQL = SQL & "	WHERE CD_GUBUN = '" & sGUBUN & "'  		"
	SQL = SQL & "	ORDER BY NO_ORDER ASC 					"
	objADO.setSql(SQL)
	ArrDetail = objADO.getArrRs()

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
                    <h2>System Code Data (DETAIL)</h2>
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
                            <strong>System Code Data (DETAIL)</strong>
                        </li>
                    </ol>
                </div>
            </div>
			
			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-12">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-search"></i> 코드 검색 <small>시스템 코드를 검색합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<div class="form-group">
									<label>코드구분 선택</label>
									<select id="sGUBUN" name="sGUBUN" class="form-control" onchange="f_reload(this.value);">
										<option value="">선택하세요.</option>
								<%
								if isarray(ArrMast) then 
									for iloop=0 to ubound(ArrMast,2)
								%>
										<option value="<%=ArrMast(0,iloop)%>" <%=getBoolean(sGUBUN=ArrMast(0,iloop),"selected","")%>><%=ArrMast(1,iloop)%></option>
								<%
									next 
								end if 
								%>
									</select>
								</div>
							</div>
						</div>



						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-list"></i> 코드 리스트</h5>
							</div>
							<div class="ibox-content">
								<div class="table-responsive">
									<form class="form-group-sm" id="frm_Code" name="frm_Code" method="post" action="codeProcessController.asp">
										<input type="hidden" id="MODE" name="MODE">
										<input type="hidden" id="RTN_STR" name="RTN_STR" value="codedetail.asp?sGUBUN=<%=sGUBUN%>">
										<input type="hidden" id="CD_GUBUN" name="CD_GUBUN" value="<%=sGUBUN%>">
										<input type="hidden" id="CD_CODE" name="CD_CODE">
										<input type="hidden" id="nm_string" name="nm_string">
							<%
							if sGUBUN <> "" then 
							%>
										<table class="table table-striped table-hover">
										<thead>
											<tr>
												<th class="text-center">코드구분</th>
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
												<td><%=sGUBUN%></td>
												<td><input class="form-control" type="text" id="CD_CODE_0" name="CD_CODE_0" size="10"></td>
												<td><input class="form-control" type="text" id="NM_CODE_0" name="NM_CODE_0" size="20"></td>
												<td><input class="form-control" type="text" id="NM_CODE1_0" name="NM_CODE1_0" size="10"></td>
												<td><input class="form-control" type="text" id="NM_CODE2_0" name="NM_CODE2_0" size="10"></td>
												<td><input class="form-control" type="text" id="NM_CODE3_0" name="NM_CODE3_0" size="10"></td>
												<td><input class="form-control" type="text" id="NM_CODE4_0" name="NM_CODE4_0" size="10"></td>
												<td><input class="form-control" type="text" id="NM_CODE5_0" name="NM_CODE5_0" size="10"></td>
												<td><input class="form-control" type="text" id="NM_CODE6_0" name="NM_CODE6_0" size="10"></td>
												<td><input class="form-control" type="text" id="NM_CODE7_0" name="NM_CODE7_0" size="10"></td>
												<td><input class="form-control text-right" type="text" id="NO_ORDER_0" name="NO_ORDER_0" size="10" value="0"></td>
												<td>
													<select name="YN_USE_0" class="form-control" >
														<option value="Y" selected="selected">예</option>
														<option value="N">아니오</option>
													</select>
												</td>
												<td>
													<button type="button" class="btn btn-success btn-sm" onclick="f_Edit('I','0','');"><i class="fa fa-plus"></i> 추가</button>
												</td>
											</tr>
											<tr>
												<th class="text-center">코드구분</th>
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
									<%If IsArray(ArrDetail) Then %>
										<%
										For iloop = 0 To UBound(ArrDetail,2) 
											CD_GUBUN = ArrDetail(0, iloop)
											CD_CODE = ArrDetail(1, iloop)
											NM_CODE = ArrDetail(2, iloop)
											NM_CODE1 = ArrDetail(3, iloop)
											NM_CODE2 = ArrDetail(4, iloop)
											NM_CODE3 = ArrDetail(5, iloop)
											NM_CODE4 = ArrDetail(6, iloop)
											NM_CODE5 = ArrDetail(7, iloop)
											NM_CODE6 = ArrDetail(8, iloop)
											NM_CODE7 = ArrDetail(9, iloop)
											NO_ORDER = ArrDetail(10, iloop)
											YN_USE = ArrDetail(11, iloop)
										%>
											<tr>
												<td><%=CD_GUBUN%></td>
												<td><input type="hidden" id="CD_CODE_<%=CD_GUBUN&CD_CODE%>" name="CD_CODE_<%=CD_GUBUN&CD_CODE%>" value="<%=CD_CODE%>"><%=CD_CODE%></td>
												<td><input type="text" class="form-control" id="NM_CODE_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE%>" size="20"></td>
												<td><input type="text" class="form-control" id="NM_CODE1_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE1_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE1%>" size="10"></td>
												<td><input type="text" class="form-control" id="NM_CODE2_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE2_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE2%>" size="10"></td>
												<td><input type="text" class="form-control" id="NM_CODE3_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE3_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE3%>" size="10"></td>
												<td><input type="text" class="form-control" id="NM_CODE4_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE4_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE4%>" size="10"></td>
												<td><input type="text" class="form-control" id="NM_CODE5_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE5_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE5%>" size="10"></td>
												<td><input type="text" class="form-control" id="NM_CODE6_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE6_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE6%>" size="10"></td>
												<td><input type="text" class="form-control" id="NM_CODE7_<%=CD_GUBUN&CD_CODE%>" name="NM_CODE7_<%=CD_GUBUN&CD_CODE%>" value="<%=NM_CODE7%>" size="10"></td>
												<td><input type="text" class="form-control text-right" id="NO_ORDER_<%=CD_GUBUN&CD_CODE%>" name="NO_ORDER_<%=CD_GUBUN&CD_CODE%>" value="<%=NO_ORDER%>" size="10"></td>
												<td>
													<select name="YN_USE_<%=CD_GUBUN&CD_CODE%>" class="form-control">
														<option value="Y" <%=getBoolean(YN_USE = "Y", "selected='selected'", "")%>>예</option>
														<option value="N" <%=getBoolean(YN_USE = "N", "selected='selected'", "")%>>아니오</option>
													</select>
												</td>
												<td>
													<button type="button" class="btn btn-info btn-sm" onclick="f_Edit('U','<%=CD_GUBUN&CD_CODE%>','<%=CD_CODE%>');"><i class="fa fa-check"></i> 수정</button>	
												</td>
											</tr>
										<%Next %>
									<%End If %>
								<%End If %>
										<tbody>
										</table>
									</form>
								</div><!-- /.table-responsive -->
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
		function f_Edit(MODE, CD_CODE, realCD_CODE)
		{
			var CD_CODE;
			var NM_CODE;
			var NO_ORDER;

			if (MODE=="I")
			{
				CD_CODE = document.getElementById("CD_CODE_0");
				NM_CODE = document.getElementById("NM_CODE_0");
				NO_ORDER = document.getElementById("NO_ORDER_0");

				if (CD_CODE.value=="") { alert("코드정보를 입력하세요."); CD_CODE.focus(); return false; }
				if (NM_CODE.value=="") { alert("코드명을 입력하세요."); NM_CODE.focus(); return false; }
				if(ValNumber(NO_ORDER, "정렬순서를", true) == false) { return false; };

				document.getElementById("CD_CODE").value = document.getElementById("CD_CODE_0").value;
				document.getElementById("nm_string").value = "0";
			}
			else if (MODE=="U")
			{
				NM_CODE = document.getElementById("NM_CODE_"+CD_CODE);
				NO_ORDER = document.getElementById("NO_ORDER_"+CD_CODE);
				if (NM_CODE.value=="") { alert("코드명을 입력하세요."); NM_CODE.focus(); return false; }
				if(ValNumber(NO_ORDER, "정렬순서를", true) == false) { return false; };

				document.getElementById("CD_CODE").value = realCD_CODE;
				document.getElementById("nm_string").value = CD_CODE;
			}
			
			document.getElementById("MODE").value = MODE;
			document.frm_Code.submit();
		}
		function f_reload(sGUBUN)
		{
			document.location.href="codedetail.asp?sGUBUN="+sGUBUN;
		}
	</script>
</body>
</html>

