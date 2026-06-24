<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_PageNavigation.asp"-->
<%
	'########################################################################################
	'#	File		: /manager/auth/UserList.asp
	'#  Create		: 조영준 / 2011.08.04
	'#	Info		: 어드민 사용자 리스트 관리
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, SQL, iloop
	Dim aryUserList, sGUBUN, sKEYWORD
	Dim CD_ADMINID
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
	SQL = SQL & " DECLARE @CNT INT														"
	SQL = SQL & " DECLARE @SNUM INT														"
	SQL = SQL & " DECLARE @ENUM INT														"
	SQL = SQL & "																		"
	SQL = SQL & " SELECT @CNT = COUNT(*)												"
	SQL = SQL & " FROM T_ADMIN															"
	'SQL = SQL & " WHERE YN_LOGIN = 'Y'													"
	SQL = SQL & " WHERE CD_ADMINID <> 'guest'											"
	SQL = SQL & "																		"
	SQL = SQL & " SET @SNUM = @CNT - ( " & pgsize & " * " & page & " ) + 1				"
	SQL = SQL & " SET @ENUM = @SNUM + " & pgsize & " - 1								"
	SQL = SQL & "																		"
	SQL = SQL & " SELECT 																"
	SQL = SQL & " 	@CNT AS CNT, *														"
	SQL = SQL & " FROM																	"
	SQL = SQL & " (																		"
	SQL = SQL & " SELECT																"
	SQL = SQL & " 	ROW_NUMBER() OVER (ORDER BY CD_ADMINID ASC) AS [No]					" '1
	SQL = SQL & " 	, CD_ADMINID,NM_ADMINPW,NM_NAME,NO_TEL,NM_EMAIL,YN_LOGIN,YN_USE		" '8
	SQL = SQL & " FROM T_ADMIN															"
	'SQL = SQL & " WHERE YN_LOGIN = 'Y'													"
	SQL = SQL & " WHERE CD_ADMINID <> 'guest'											"
	SQL = SQL & " ) AS L																"
	SQL = SQL & " WHERE No BETWEEN @SNUM AND @ENUM										"
	SQL = SQL & " ORDER BY																"
	SQL = SQL & " 	No DESC																"

	objADO.setSql(SQL)
	aryUserList = objADO.getArrRs()

	If IsArray(aryUserList) Then
	'If Not RS.Eof then
		tn =  aryUserList(0,0)								'총 페이지수
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
</head>
<body class="skin-2">
	<div id="wrapper">
        <!--#include virtual="/manager/inc/inc.ui.navi.asp"-->
		
		<div id="page-wrapper" class="gray-bg">
			
			<!--#include virtual="/manager/inc/inc.ui.top.asp"-->

			<div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-10">
                    <h2>접속계정 관리</h2>
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
                            <strong>접속계정 관리</strong>
                        </li>
                    </ol>
                </div>
            </div>
            
			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-12">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-user"></i> 접속계정 관리 <small>관리자 계정의 정보를 관리합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<div class="table-responsive">
									<form class="" id="frm_UserList" name="frm_UserList">
										<input type="hidden" id="MODE" name="MODE" value=""/>
										<input type="hidden" id="update_CD_ADMINID" name="update_CD_ADMINID" value="" />

										<table class="table table-striped table-hover">
											<thead>
												<tr>
													<th class="text-center">추가</th>
													<th class="text-center">ID</th>
													<th class="text-center">비밀번호</th>
													<th class="text-center">이름</th>
													<th class="text-center">연락처</th>
													<th class="text-center">이메일</th>
													<th class="text-center">어드민로그인여부</th>
													<th class="text-center">사용여부</th>
													<th class="text-center">관리</th>
												</tr>
											</thead>
											<tbody>
												<tr>
													<td class="text-center">NEW</td>
													<td class="text-center"><input type="text" class="form-control input-sm" id="CD_ADMINID" name="CD_ADMINID" /></td>
													<td class="text-center"><input type="password" class="form-control input-sm" id="NM_ADMINPW" name="NM_ADMINPW" /></td>
													<td class="text-center"><input type="text" class="form-control input-sm" id="NM_NAME" name="NM_NAME" /></td>
													<td class="text-center"><input type="text" class="form-control input-sm" id="NO_TEL" name="NO_TEL" /></td>
													<td class="text-center"><input type="text" class="form-control input-sm" id="NM_EMAIL" name="NM_EMAIL" /></td>
													<td class="text-center">
														<select class="form-control input-sm" id="YN_LOGIN" name="YN_LOGIN">
															<option value="Y" >예</option>
															<option value="N" >아니오</option>
														</select>
													</td>
													<td class="text-center">
														<select class="form-control input-sm" id="YN_USE" name="YN_USE">
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

										
									<%
									If IsArray(aryUserList) Then
										response.write tn & " 명의 관리자가 있습니다."
									end if 
									%>
										<table class="table table-striped table-hover">
											<thead>
												<tr>
													<th class="text-center">순번</th>
													<th class="text-center">ID</th>
													<th class="text-center">비밀번호</th>
													<th class="text-center">이름</th>
													<th class="text-center">연락처</th>
													<th class="text-center">이메일</th>
													<th class="text-center">어드민로그인여부</th>
													<th class="text-center">사용여부</th>
													<th class="text-center">관리</th>
												</tr>
											</thead>
											<tbody>
											<%
												If IsArray(aryUserList) Then
												For iloop = 0 To UBound(aryUserList,2)
													CD_ADMINID = aryUserList(2,iloop)
											%>
												<tr>
													<td class="text-center"><%=aryUserList(1,iloop)%></td>
													<td class="text-center"><%=CD_ADMINID%></td>
													<td class="text-center">
														<input type="password" class="form-control input-sm" id="NM_ADMINPW_<%=CD_ADMINID%>" name="NM_ADMINPW_<%=CD_ADMINID%>" value="<%=aryUserList(3,iloop)%>" />
													</td>
													<td class="text-center">
														<input type="text" class="form-control input-sm" id="NM_NAME_<%=CD_ADMINID%>" name="NM_NAME_<%=CD_ADMINID%>" value="<%=aryUserList(4,iloop)%>" />
													</td>
													<td class="text-center">
														<input type="text" class="form-control input-sm" id="NO_TEL_<%=CD_ADMINID%>" name="NO_TEL_<%=CD_ADMINID%>" value="<%=aryUserList(5,iloop)%>" />
													</td>
													<td class="text-center">
														<input type="text" class="form-control input-sm" id="NM_EMAIL_<%=CD_ADMINID%>" name="NM_EMAIL_<%=CD_ADMINID%>" value="<%=aryUserList(6,iloop)%>" />
													</td>
													<td class="text-center">
														<select class="form-control input-sm" id="YN_LOGIN_<%=CD_ADMINID%>" name="YN_LOGIN_<%=CD_ADMINID%>" >
															<option value="Y" <%=getBoolean(aryUserList(7,iloop)="Y","selected","")%>>예</option>
															<option value="N" <%=getBoolean(aryUserList(7,iloop)="N","selected","")%>>아니오</option>
														</select>
													</td>
													<td class="text-center">
														<select class="form-control input-sm" id="YN_USE_<%=CD_ADMINID%>" name="YN_USE_<%=CD_ADMINID%>" >
															<option value="Y" <%=getBoolean(aryUserList(8,iloop)="Y","selected","")%>>예</option>
															<option value="N" <%=getBoolean(aryUserList(8,iloop)="N","selected","")%>>아니오</option>
														</select>
													</td>
													<td class="text-center">
														<button type="button" class="btn btn-info btn-sm" onclick="fn_userList('U','<%=CD_ADMINID%>');"><i class="fa fa-edit"></i> 수정</button>
													</td>
												</tr>
											<%
												Next
											Else 
											%>
												<tr>
													<td class="text-center" colspan="9">데이터가 존재하지 않습니다.</td>
												</tr>
											<%
											End If 
											%>
											</tbody>
											<tfoot>
												<tr>
													<td class="text-center" colspan="9">
														<%
														If IsArray(aryUserList) Then
															Call PageNavigation("/manager/auth/UserList.asp","", tn, pgcount, startpage, page, setsize, prevpage, nextpage)
														End If
														%>
													</td>
												</tr>
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
			document.frm_UserList.action = "UserList_proc.asp";
			document.frm_UserList.submit();
		}

		function fn_userList(Mode,ID)
		{
			var MODE	   = document.getElementById("MODE");
			var CD_ADMINID = document.getElementById("CD_ADMINID");
			var NM_NAME	   = document.getElementById("NM_NAME");
			var update_CD_ADMINID = document.getElementById("update_CD_ADMINID");

			MODE.value = Mode;

			if (Mode == "I")
			{	
				if (CD_ADMINID.value == "")
				{
					alert("ID를 입력해주세요.");
					CD_ADMINID.focus();
					return false;
				}
				if (NM_NAME.value == "")
				{
					alert("이름를 입력해주세요.");
					NM_NAME.focus();
					return false;
				}	
				fn_submit();
			}
			else if (Mode == "U")
			{
				update_CD_ADMINID.value = ID;
				fn_submit();
			}			
		}
	</script>
	<iframe type="hiddenframe" id="hd_frame" name="hd_frame" style="display:none;"></iframe>	
</body>
</html>