<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/auth/UserList.asp
	'#  Create		: domoyosi / 2015.12
	'#	Info		: 관리자 비밀번호 수정
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, Rs, SQL, iloop

	Set objADO = new clsADO
		
	SQL = ""
	SQL = SQL & " SELECT																"
	SQL = SQL & " 	CD_ADMINID,NM_ADMINPW,NM_NAME,NO_TEL,NM_EMAIL,YN_LOGIN,YN_USE		"
	SQL = SQL & " FROM T_ADMIN															"
	SQL = SQL & " WHERE CD_ADMINID = '" & m_UserId & "'									"

	objADO.setSql(SQL)
	Set Rs = objADO.getRs()	
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
                    <h2>비밀번호 수정</h2>
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
                            <strong>비밀번호 수정</strong>
                        </li>
                    </ol>
                </div>
            </div>
			
			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-4">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-key"></i> 비밀번호 수정 <small>사용자의 비밀번호를 수정합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<form id="frm_UserInfo" name="frm_UserInfo">
								<%
									If Not Rs.EOF Then
								%>
									<div class="form-group">
										<label>기존 비밀번호 *</label>
										<input type="password" class="form-control input-sm" id="NM_ADMINPW" name="NM_ADMINPW" tabindex="1" placeholder="기존 비밀번호를 입력해주세요." />
									</div>

									<div class="form-group">
										<label>새로운 비밀번호 *</label>
										<input type="password" class="form-control input-sm" id="NM_ADMINPW1" name="NM_ADMINPW1" tabindex="2" placeholder="새로운 비밀번호를 입력해주세요." />
									</div>

									<div class="form-group">
										<label>비밀번호 재입력 *</label>
										<input type="password" class="form-control input-sm" id="NM_ADMINPW2" name="NM_ADMINPW2" tabindex="3" placeholder="새로운 비밀번호를 재입력해주세요." />
									</div>

									<div class="form-group">
										<button type="button" class="btn btn-info btn-sm" onclick="fn_ok();"><i class="fa fa-check"></i> 수정</button>	
										<button type="button" class="btn btn-danger btn-sm" onclick="fn_cancle();"><i class="fa fa-refresh"></i> 취소</button>	
									</div>

								<%
									End If 

									Set objADO = Nothing
								%>
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

	<script type="text/javascript">
		
		$(document).ready(function() {
			$("#NM_ADMINPW").focus();
		});

		function fn_ok() {
			var NM_ADMINPW	= $("#NM_ADMINPW");
			var NM_ADMINPW1 = $("#NM_ADMINPW1");
			var NM_ADMINPW2	= $("#NM_ADMINPW2");

			if (NM_ADMINPW.val() == "")
			{
				alert("기존 비밀번호를 입력해주세요.");
				NM_ADMINPW.focus();
				return false;
			}
			if (NM_ADMINPW1.val() == "")
			{
				alert("새로운 비밀번호를 입력해주세요.");
				NM_ADMINPW1.focus();
				return false;
			}
			if (NM_ADMINPW2.val() == "")
			{
				alert("비밀번호를 한번더 입력해주세요.");
				NM_ADMINPW2.focus();
				return false;
			}
			if (NM_ADMINPW1.val() != NM_ADMINPW2.val())
			{
				alert("새로운 비밀번호가 서로 틀립니다.");
				NM_ADMINPW1.select();
				return false;
			}
			document.frm_UserInfo.target = "hd_frame";
			document.frm_UserInfo.method = "post";
			document.frm_UserInfo.action = "UserInfo_proc.asp";
			document.frm_UserInfo.submit();
		}

		function fn_cancle() {
			$("#NM_ADMINPW").val("");
			$("#NM_ADMINPW1").val("");
			$("#NM_ADMINPW2").val("");
		}
	</script>
	<iframe type="hiddenframe" id="hd_frame" name="hd_frame" style="display:none;"></iframe>
</body>
</html>



