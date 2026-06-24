<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp"-->
<%
	'########################################################################################
	'#	File		: /manager/email/receive_email.asp
	'#  Create		: 
	'#	Info		: 이메일 수신 관리
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim SQL, iloop, objADO
	Dim CD_ADMINID
	Dim aryUserList, sGUBUN, sKEYWORD

	Set objADO = new clsADO
		
	SQL = ""
	SQL = SQL & "SELECT													"
	SQL = SQL & "	CD_GUBUN,CD_CODE,NM_CODE,NM_CODE1,NM_CODE2,NM_CODE3		"	'5
	SQL = SQL & "	,NM_CODE4,NM_CODE5,NM_CODE6,NM_CODE7,NO_ORDER			"	'10
	SQL = SQL & "	,YN_USE													"
	SQL = SQL & "FROM														"
	SQL = SQL & "	T_CODE													"
	SQL = SQL & "WHERE													"
	SQL = SQL & "	CD_GUBUN = 'COUNSEL_MAIL_RECV'		"

	SQL = SQL & "ORDER BY													"
	SQL = SQL & "	NM_CODE1 ASC, NM_CODE2 ASC, NM_CODE3 ASC, NO_ORDER ASC											"

	objADO.setSql(SQL)
	aryUserList = objADO.getArrRs()

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
                    <h2>상담 메일 수신자 관리</h2>
                    <ol class="breadcrumb">
                        <li>
                            <a href="/manager/">Home</a>
                        </li>
                        <li>
                            <a href="javascript:void(0)">시스템 관리</a>
                        </li>
						<li>
                            <a href="javascript:void(0)">사용자관리</a>
                        </li>
                        <li class="active">
                            <strong>상담 메일 수신자 관리</strong>
                        </li>
                    </ol>
                </div>
            </div>
            
			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-12">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-user"></i> 상담 메일 수신자 관리 <small>상담글 등록시 메일 알림을 수신할 이메일 주소를 관리합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<div class="table-responsive">
									<form class="" id="frm_UserList" name="frm_UserList" autocomplete="off" >
										<input type="hidden" id="MODE" name="MODE" value=""/>
										<input type="hidden" id="update_CD_CODE" name="update_CD_CODE" value=""/>
										
										<table class="table table-striped table-hover">
											<thead>
												<tr>
													<th class="text-center">추가</th>
													<th class="text-center">상담게시판</th>
													<th class="text-center">이름</th>
													<th class="text-center">이메일</th>
													<th class="text-center">사용여부</th>
													<th class="text-center">관리</th>
												</tr>
											</thead>
											<tbody>
												<tr>
													<td class="text-center">NEW</td>
													<td class="text-center">
														<select class="form-control" id="NM_CODE1" name="NM_CODE1">
															<option value="1003" >온라인상담</option>
															<!--<option value="1008" >빠른상담</option>-->
														</select>
													</td>
													<td class="text-center"><input type="text" class="form-control" id="NM_CODE2" name="NM_CODE2" autocomplete="off" /></td>
													<td class="text-center"><input type="text" class="form-control" id="NM_CODE3" name="NM_CODE3" autocomplete="off" placeholder="email_address@domain" /></td>
													<td class="text-center">
														<select class="form-control" id="YN_USE" name="YN_USE">
															<option value="Y" >예</option>
															<option value="N" >아니오</option>
														</select>
													</td>
													<td class="text-center">
														<button type="button" class="btn btn-success btn-sm" onclick="fn_userList('I','');"><i class="fa fa-plus"></i> 추가</button>
													</td>
												</tr>
											</tbody>
										</table>

										<table class="table table-striped table-hover">
											<thead>
												<tr>
													<th class="text-center">상담게시판</th>
													<th class="text-center">이름</th>
													<th class="text-center">이메일</th>
													<th class="text-center">사용여부</th>
													<th class="text-center">관리</th>
												</tr>
											</thead>
											<tbody>
											<%
												If IsArray(aryUserList) Then
												For iloop = 0 To UBound(aryUserList,2)
											%>
												<tr>
													<td class="text-center">
														<select class="form-control" id="NM_CODE1_<%=aryUserList(1,iloop)%>" name="NM_CODE1_<%=aryUserList(1,iloop)%>">
															<option value="1003" <%=getBoolean(aryUserList(3,iloop) = "1003", "selected='selected'", "")%>>온라인상담</option>
															<!--<option value="1008" <%=getBoolean(aryUserList(3,iloop) = "1008", "selected='selected'", "")%>>빠른상담</option>-->
														</select>
													</td>
													<td class="text-center">
														<input type="text" class="form-control" id="NM_CODE2_<%=aryUserList(1,iloop)%>" name="NM_CODE2_<%=aryUserList(1,iloop)%>" value="<%=aryUserList(4,iloop)%>" autocomplete="off" />
													</td>
													<td class="text-center">
														<input type="text" class="form-control" id="NM_CODE3_<%=aryUserList(1,iloop)%>" name="NM_CODE3_<%=aryUserList(1,iloop)%>" value="<%=aryUserList(5,iloop)%>" autocomplete="off" />
													</td>
													<td class="text-center">
														<select class="form-control input-sm" id="YN_USE_<%=aryUserList(1,iloop)%>" name="YN_USE_<%=aryUserList(1,iloop)%>" >
															<option value="Y" <%=getBoolean(aryUserList(11,iloop)="Y","selected","")%>>예</option>
															<option value="N" <%=getBoolean(aryUserList(11,iloop)="N","selected","")%>>아니오</option>
														</select>
													</td>
													<td class="text-center">
														<button type="button" class="btn btn-info btn-sm" onclick="fn_userList('U','<%=aryUserList(1,iloop)%>');"><i class="fa fa-edit"></i> 수정</button>
														<button type="button" class="btn btn-danger btn-sm" onclick="fn_userList('D','<%=aryUserList(1,iloop)%>');"><i class="fa fa-trash"></i> 삭제</button>
													</td>
												</tr>
											<%
												Next
											Else 
											%>
												<tr>
													<td class="text-center" colspan="5">데이터가 존재하지 않습니다.</td>
												</tr>
											<%
											End If 
											%>
											</tbody>
											<tfoot>
												
											</tfoot>
										</table>
									</form>
								</div><!-- /.table-responsive -->


							</div>
						</div>
					
					</div>

                </div>
                <!-- /.col-lg-12 -->
            </div>
            <!-- /.row -->
		</div>
		<!-- /#page-wrapper -->
	</div>
	<!-- /#wrapper -->

	<script type="text/javascript">
		function fn_submit()
		{
			document.frm_UserList.target = "hd_frame";
			document.frm_UserList.method = "post";
			document.frm_UserList.action = "receive_email_proc.asp";
			document.frm_UserList.submit();
		}

		function fn_userList(Mode,ID)
		{
			var MODE		= document.getElementById("MODE");
			var NM_CODE2	= document.getElementById("NM_CODE2");
			var NM_CODE3	= document.getElementById("NM_CODE3");
			var update_CD_CODE = document.getElementById("update_CD_CODE");

			MODE.value = Mode;

			if (Mode == "I")
			{	
				if (NM_CODE2.value == "")
				{
					alert("이름을 입력해주세요.");
					NM_CODE2.focus();
					return false;
				}
				if (NM_CODE3.value == "")
				{
					alert("이메일을 입력해주세요.");
					NM_CODE3.focus();
					return false;
				}

				if (ValEmail_V3(NM_CODE3.value, '', true) == false)
				{
					alert("유효하지 않은 이메일 형식 입니다.");
					NM_CODE3.focus();
					return false;
				}

				fn_submit();
			}
			else if (Mode == "U")
			{
				update_CD_CODE.value = ID;
				
				var update_NM_CODE2	= document.getElementById("NM_CODE2_"+ID);
				var update_NM_CODE3	= document.getElementById("NM_CODE3_"+ID);

				if (update_NM_CODE2.value == "")
				{
					alert("이름을 입력해주세요.");
					update_NM_CODE2.focus();
					return false;
				}
				if (update_NM_CODE3.value == "")
				{
					alert("이메일을 입력해주세요.");
					update_NM_CODE3.focus();
					return false;
				}

				if (ValEmail_V3(update_NM_CODE3.value, '', true) == false)
				{
					alert("유효하지 않은 이메일 형식 입니다.");
					update_NM_CODE3.focus();
					return false;
				}

				fn_submit();
			}
			else if (Mode == "D")
			{
				update_CD_CODE.value = ID;
				
				if (confirm("삭제 하시겠습니까?"))
				{
					fn_submit();
				}

			}
		}
	</script>
	<iframe type="hiddenframe" id="hd_frame" name="hd_frame" style="display:none;"></iframe>	
</body>
</html>