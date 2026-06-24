<%
	Dim menu_objADO, menu_SQL, menu_ploop, menu_iloop, menu_jloop, menu_kloop
	Dim menu_CD_FORM_0, menu_CD_PFORM_0, menu_CD_ROOT_0, menu_NO_DEPTH_0, menu_CD_PART_0, menu_NM_CODE_0, menu_NM_FORM_0, menu_NM_URL_0, menu_CD_EDITION_0, menu_YN_USER_0, menu_YN_USE_0, menu_CD_TARGET_0
	Dim menu_CD_FORM_1, menu_CD_PFORM_1, menu_CD_ROOT_1, menu_NO_DEPTH_1, menu_CD_PART_1, menu_NM_CODE_1, menu_NM_FORM_1, menu_NM_URL_1, menu_CD_EDITION_1, menu_YN_USER_1, menu_YN_USE_1, menu_CD_TARGET_1
	Dim menu_CD_FORM_2, menu_CD_PFORM_2, menu_CD_ROOT_2, menu_NO_DEPTH_2, menu_CD_PART_2, menu_NM_CODE_2, menu_NM_FORM_2, menu_NM_URL_2, menu_CD_EDITION_2, menu_YN_USER_2, menu_YN_USE_2, menu_CD_TARGET_2
	Dim menu_ArrMenuList
'	Dim profile_SQL
'	Dim profileList

	Set menu_objADO = new clsADO
	menu_objADO.setConString(m_DB)

	menu_SQL = ""
	menu_SQL = menu_SQL & " SELECT																				"
	menu_SQL = menu_SQL & "	PG.CD_FORM, PG.CD_PFORM, ISNULL(RT.CD_PFORM,'') AS CD_ROOT, PG.NO_DEPTH, PG.CD_PART	"	'4
	menu_SQL = menu_SQL & " 	, PG.NM_FORM, PG.NM_URL, PG.CD_NAMESPACE											"	'7
	menu_SQL = menu_SQL & " 	, PG.CD_EDITION, PG.NO_ORDER, PG.MM_HELP, PG.YN_USER, PG.YN_USE						"	'12
	menu_SQL = menu_SQL & " 	, CASE WHEN AU.CD_ADMINID IS NULL THEN 'N' ELSE 'Y' END AS [YN_USEDMENU]			"	'13 (메뉴 사용여부)
	menu_SQL = menu_SQL & "	, PG.CD_TARGET																		"	'14
	menu_SQL = menu_SQL & " FROM T_MENU AS PG																		"
	menu_SQL = menu_SQL & " 		LEFT OUTER JOIN																	"
	menu_SQL = menu_SQL & " 		(																				"
	menu_SQL = menu_SQL & " 			T_AUTH AS AU																"
	menu_SQL = menu_SQL & " 			INNER JOIN T_ADMIN AS MB													"
	menu_SQL = menu_SQL & " 			ON MB.YN_LOGIN = 'Y' AND MB.YN_USE = 'Y' AND MB.CD_ADMINID = '" & m_UserID & "' AND MB.CD_ADMINID = AU.CD_ADMINID	"
	menu_SQL = menu_SQL & " 		) ON AU.CD_FORM = PG.CD_FORM AND PG.YN_USE = 'Y'													"
	menu_SQL = menu_SQL & " 		LEFT OUTER JOIN T_MENU AS RT													"
	menu_SQL = menu_SQL & " 	ON RT.NO_DEPTH IN (0,1) AND RT.YN_USE = 'Y' AND RT.CD_FORM = PG.CD_PFORM 			"
	menu_SQL = menu_SQL & "	WHERE AU.CD_ADMINID IS NOT NULL																"
	menu_SQL = menu_SQL & " ORDER BY PG.NO_ORDER ASC																"

	menu_objADO.setSql(menu_SQL)
	menu_ArrMenuList = menu_objADO.getArrRs()


'	profile_SQL = ""
'	profile_SQL = profile_SQL & "SELECT 	"
'	profile_SQL = profile_SQL & " 	NM_PROFILE_IMAGE "
'	profile_SQL = profile_SQL & " FROM T_ADMIN WHERE CD_ADMINID = '" & m_UserID & "'	"

'	menu_objADO.setSql(profile_SQL)
'	profileList = menu_objADO.getArrRs()

	Set menu_objADO = Nothing

	Dim profileImage

'	If IsArray(profileList) Then
'		profileImage = profileList(0, 0)
'	End If

	Dim menu_active_check_url : menu_active_check_url = Request.ServerVariables("URL")
	Dim menu_active_check_querystring : menu_active_check_querystring = Request.ServerVariables("QUERY_STRING")

	'Response.write menu_active_check_url
	'Response.write menu_active_check_url & "?" & menu_active_check_querystring
%>


	<nav class="navbar-default navbar-static-side" role="navigation">
        <div class="sidebar-collapse">
            <ul class="nav metismenu" id="side-menu">
                <li class="nav-header">
                    <div class="dropdown profile-element">

                        <a data-toggle="dropdown" class="dropdown-toggle" href="#">
                        <span class="clear"> <span class="block m-t-xs"> <strong class="font-bold"><%=m_UserNM%></strong>
                         </span> <span class="text-muted text-xs block"><%=m_UserID%> <b class="caret"></b></span> </span> </a>
                        <ul class="dropdown-menu animated fadeInRight m-t-xs">
							<!--
                            <li><a href="profile.html">Profile</a></li>
                            <li><a href="contacts.html">Contacts</a></li>
                            <li><a href="mailbox.html">Mailbox</a></li>
                            <li class="divider"></li>
							-->
                            <li><a href="/manager/logout.asp">Logout</a></li>
                        </ul>
                    </div>
                    <div class="logo-element">
                        Admin
                    </div>
                </li>

				<%
				If IsArray(menu_ArrMenuList) Then
					For menu_iloop = 0 To UBound(menu_ArrMenuList,2)
						menu_CD_FORM_0	= menu_ArrMenuList(0,menu_iloop)
						menu_CD_PFORM_0	= menu_ArrMenuList(1,menu_iloop)
						menu_NO_DEPTH_0	= menu_ArrMenuList(3,menu_iloop)
						menu_CD_PART_0	= menu_ArrMenuList(4,menu_iloop)
						menu_NM_FORM_0	= menu_ArrMenuList(5,menu_iloop)
						menu_NM_URL_0	= menu_ArrMenuList(6,menu_iloop)
						menu_CD_EDITION_0 = menu_ArrMenuList(8,menu_iloop)
						menu_CD_TARGET_0	= menu_ArrMenuList(14,menu_iloop)
						If menu_NO_DEPTH_0 = 0 Then '0레벨
				%>
				<li class="manager-app-menu-level-0" data-codeform0="<%=menu_CD_FORM_0%>">

				<%If menu_NM_URL_0 <> "" Then %>
					<a href="<%=menu_NM_URL_0%>" target="_<%=menu_CD_TARGET_0%>" class=""><i class="fa <%=menu_CD_EDITION_0%>"></i> <span><%=menu_NM_FORM_0%></span></a>
				<%Else %>
					<a href="#"><i class="fa <%=menu_CD_EDITION_0%>"></i> <span class="nav-label"><%=menu_NM_FORM_0%> </span><span class="fa arrow"></span></a>
					<ul class="nav nav-second-level collapse">
							<%
							For menu_jloop = 0 To UBound(menu_ArrMenuList,2)
								menu_CD_FORM_1	= menu_ArrMenuList(0,menu_jloop)
								menu_CD_PFORM_1	= menu_ArrMenuList(1,menu_jloop)
								menu_NO_DEPTH_1	= menu_ArrMenuList(3,menu_jloop)
								menu_NM_FORM_1	= menu_ArrMenuList(5,menu_jloop)
								If menu_NO_DEPTH_1 = 1 And menu_CD_PFORM_1 = menu_CD_FORM_0 Then '1레벨
							%>
						<li class="manager-app-menu-level-1">
							<a href="javascript:void(0);"><span class="text-success"><%=menu_NM_FORM_1%></span></a>
						</li>
							<%
								For menu_kloop = 0 To UBound(menu_ArrMenuList,2)
									menu_CD_FORM_2	= menu_ArrMenuList(0,menu_kloop)
									menu_CD_PFORM_2	= menu_ArrMenuList(1,menu_kloop)
									menu_NO_DEPTH_2	= menu_ArrMenuList(3,menu_kloop)
									menu_NM_FORM_2	= menu_ArrMenuList(5,menu_kloop)
									menu_NM_URL_2	= menu_ArrMenuList(6,menu_kloop)
									menu_CD_EDITION_2	= menu_ArrMenuList(8,menu_kloop)
									menu_CD_TARGET_2	= menu_ArrMenuList(14,menu_kloop)
									If menu_NO_DEPTH_2 = 2 And menu_CD_PFORM_2 = menu_CD_FORM_1 Then '2레벨
							%>
						<li class="manager-app-menu-level-2" data-codeform0="<%=menu_CD_FORM_0%>" data-codeform2="<%=menu_CD_FORM_2%>">
							<a href="<%=menu_NM_URL_2%>" target="_<%=menu_CD_TARGET_2%>" class=""><i class="fa <%=menu_CD_EDITION_2%>"></i> <%=menu_NM_FORM_2%></a>
						</li>
							<%
									End If
								Next
								End If
							Next
							%>
					</ul>
				<%End If %>


				</li>
				<%
						End If
					Next
				End If
				%>
            </ul>

        </div>
    </nav>

		<script type="text/javascript">

		$(document).ready(function() {
			$height = $("#page-wrapper").height();
			$(".navbar-default").height($height);


			$("li.manager-app-menu-level-0").click(function(e) {
				$.cookie("admin_menu_code_form_0", $(this).data("codeform0"), { expires: 1, path: "/", domain: "<%=CookieDomain%>", secure: false });
				$.cookie("admin_menu_code_form_2", "", { path: "/", domain: "<%=CookieDomain%>", secure: false });
				//console.log(0);
				e.stopPropagation();	// 부모 li 에 클릭 버블링 방지
			});

			$("li.manager-app-menu-level-2").click(function(e) {
				//alert($(this).data("codeform0"));
				//$(this).data("codeform0");
				//$(this).data("codeform2");
				$.cookie("admin_menu_code_form_0", $(this).data("codeform0"), { expires: 1, path: "/", domain: "<%=CookieDomain%>", secure: false });
				$.cookie("admin_menu_code_form_2", $(this).data("codeform2"), { expires: 1, path: "/", domain: "<%=CookieDomain%>", secure: false });
				//console.log(2);
				e.stopPropagation();	// 부모 li 에 클릭 버블링 방지
			});
		});

		$(window).load(function() {

			$("li.manager-app-menu-level-0").removeClass("active");
			$("li.manager-app-menu-level-2").removeClass("active");

			var admin_menu_code_form_0 = "";
			var admin_menu_code_form_2 = "";

			if (($.cookie("admin_menu_code_form_0") !== undefined) && ($.cookie("admin_menu_code_form_0") != "")) {
				admin_menu_code_form_0 = $.cookie("admin_menu_code_form_0");
			}

			//console.log("$.cookie admin_menu_code_form_2: " + $.cookie("admin_menu_code_form_2"));

			if (($.cookie("admin_menu_code_form_2") !== undefined) && ($.cookie("admin_menu_code_form_2") != "")) {
				admin_menu_code_form_2 = $.cookie("admin_menu_code_form_2");
			}

			//console.log("admin_menu_code_form_0: " + admin_menu_code_form_0);
			//console.log("admin_menu_code_form_2: " + admin_menu_code_form_2);

			// 0 depth 를 선택한 경우 활성화 처리
			if (admin_menu_code_form_0 != "" && admin_menu_code_form_2 == "") {

				$("li.manager-app-menu-level-0").each(function() {
					if ($(this).data("codeform0") == admin_menu_code_form_0) {
						$(this).addClass("active");
						//$(this).parent("ul").prev("a").click();

						//if ($(this).find("a").attr("href").indexOf("<%=menu_active_check_url%>") == -1) {
						//	document.location.href=$(this).find("a").attr("href");
						//}

						return false;	// break;
					}
				});
			}

			// 2 depth 를 선택한 경우 활성화 처리
			if (admin_menu_code_form_0 != "" && admin_menu_code_form_2 != "") {

				$("li.manager-app-menu-level-2").each(function() {
					if ($(this).data("codeform0") == admin_menu_code_form_0 && $(this).data("codeform2") == admin_menu_code_form_2) {
						$(this).addClass("active");
						//$(this).parent("ul").prev("a").click();
						$(this).parent("ul").parent("li").addClass("active");
						$(this).parent("ul").addClass("collapse in");
						//if ($(this).find("a").attr("href").indexOf("<%=menu_active_check_url%>") == -1) {
						//	document.location.href=$(this).find("a").attr("href");
						//}

						return false;	// break;
					}
				});
			}
		});

		</script>
