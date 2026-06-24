<!-- #include virtual = "/Include/Config.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/login.asp
	'#  Create		: domoyosi / 2015.12
	'#	Info		: 어드민 로그인페이지
	'#	Update		:
	'#	Update Memo	:
	'########################################################################################
%>
<!doctype html>
<html>
<head>
	<!--#include virtual="/manager/inc/inc.ui.head.asp"-->
</head>
<body class="skin-2 gray-bg" style="background: url('/manager/images/nav_bg.gif')">

    <div class="middle-box text-center loginscreen animated fadeInDown">
        <div>
			<div>
                <h1 class="logo-name">
                	<!-- logo -->
                	<img src="/html/images/foot_logo.png" style="" />
                </h1>
            </div>

			<h3 class="text-white">Administrator</h3>

            <form class="m-t" role="form" name="frmLogin" autocomplete="off">
                <div class="form-group">
                    <input class="form-control" placeholder="ID" id="userid" name="userid" type="text" autocomplete="off" />
                </div>
                <div class="form-group">
                    <input class="form-control" placeholder="Password" id="password" name="password" type="password" value="" autocomplete="off" />
                </div>
				<div class="form-group text-left">
					<label class="checkbox-inline"> <input id="id_save" name="id_save" type="checkbox" checked="checked" /> <span class="text-white">아이디 저장</span> </label>
				</div>
                <button type="button" class="btn btn-success block full-width m-b" onclick="fn_logincheck()">Login</button>

            </form>
			<p class="m-t"> <small class="text-white">&copy;Peter's</small> </p>
        </div>
    </div>

<script type="text/javascript">
	$(document).ready(function() {

		// 메뉴 활성화 쿠키 초기화
		$.removeCookie("admin_menu_code_form_0", { path: "/", domain: "<%=CookieDomain%>", secure: false });
		$.removeCookie("admin_menu_code_form_2", { path: "/", domain: "<%=CookieDomain%>", secure: false });

		//$.cookie("layer_popup_20150317_001", "close", { expires: 365, path: "/", domain: "<%=CookieDomain%>", secure: false });

		if (($.cookie("admin_save_id") === undefined) || ($.cookie("admin_save_id") == "")) {
			$("#userid").focus();
		}

		if (($.cookie("admin_save_id") !== undefined) && ($.cookie("admin_save_id") != "")) {
			$("#userid").val( $.cookie("admin_save_id") );
			$("#password").focus();
		}

		$("#password").keyup(function(e) {
			if (e.keyCode == 13) {
				fn_logincheck();
			}
		});
	});

	function fn_logincheck() {
		var id = $("#userid");
		var pw = $("#password");
		var id_save = $("#id_save");

		if ($.trim(id.val()) == "") {
			alert("아이디를 입력하세요.");
			id.focus();
			return false;
		}

		if ($.trim(pw.val()) == "") {
			alert("비밀번호를 입력하세요.");
			pw.focus();
			return false;
		}

		// 아이디 기억하기
		if (id_save.prop("checked") == true) {
			$.cookie("admin_save_id", id.val(), { expires: 365, path: "/", domain: "<%=CookieDomain%>", secure: false });
		} else {
			$.removeCookie("admin_save_id", { path: "/", domain: "<%=CookieDomain%>", secure: false });
		}

		document.frmLogin.method = "post";
		document.frmLogin.action = "/manager/loginIEController.asp"
		document.frmLogin.submit();
	}
</script>
<!-- #include virtual = "/manager/inc/inc.ui.hidden.asp" -->
</body>
</html>
