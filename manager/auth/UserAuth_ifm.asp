<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/auth/adminList_ifm.asp
	'#  Create		: domoyosi / 2015
	'#	Info		: 관리자별 사용 프로그램 관리
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, Rs, SQL, iloop, jloop, kloop, ploop
	Dim ArrMenuList
	Dim CD_ADMINID, NM_NAME
	Dim CD_FORM, CD_PFORM, NM_FORM, CD_PARENT 
	Dim NO_DEPTH, FORMPART, YN_USEDMENU, CD_ROOT

	CD_ADMINID = ReqQ("CD_ADMINID")
	'CD_FORM = ReqQ("CD_FORM")
	NM_NAME = ReqQ("NM_NAME")
	
	'response.write CD_ADMINID & "<br>"
	'response.write CD_FORM & "<br>"
	'response.write NM_NAME & "<br>"
	'response.End

	Set objADO = new clsADO

	SQL = ""
	SQL = SQL & " SELECT																				" 
	SQL = SQL & "	PG.CD_FORM, PG.CD_PFORM, ISNULL(RT.CD_PFORM,'') AS CD_ROOT, PG.NO_DEPTH, PG.CD_PART	"	'4
	SQL = SQL & " 	, PG.NM_FORM, PG.NM_URL, PG.CD_NAMESPACE											"	'7
	SQL = SQL & " 	, PG.CD_EDITION, PG.NO_ORDER, PG.MM_HELP, PG.YN_USER, PG.YN_USE						"	'12
	SQL = SQL & " 	, CASE WHEN AU.CD_ADMINID IS NULL THEN 'N' ELSE 'Y' END AS [YN_USEDMENU]			"	'13 (메뉴 사용여부)
	SQL = SQL & " FROM T_MENU AS PG																		"
	SQL = SQL & " 		LEFT OUTER JOIN																	"
	SQL = SQL & " 		(																				"
	SQL = SQL & " 			T_AUTH AS AU																"
	SQL = SQL & " 			INNER JOIN T_ADMIN AS MB													"
	SQL = SQL & " 			ON MB.YN_LOGIN = 'Y' AND MB.YN_USE = 'Y' AND MB.CD_ADMINID = '" & CD_ADMINID & "' AND MB.CD_ADMINID = AU.CD_ADMINID	"
	SQL = SQL & " 		) ON AU.CD_FORM = PG.CD_FORM													"
	SQL = SQL & " 		LEFT OUTER JOIN T_MENU AS RT													"																
	SQL = SQL & " 	ON RT.NO_DEPTH IN (0,1) AND RT.YN_USE = 'Y' AND RT.CD_FORM = PG.CD_PFORM 			"
	SQL = SQL & " WHERE PG.YN_USER = 'Y' AND PG.YN_USE = 'Y'											"
	SQL = SQL & " ORDER BY PG.NO_ORDER ASC																"

	objADO.setSql(SQL)
	ArrMenuList = objADO.getArrRs()
%>
<!doctype html>
<html>
<head>
	<!--#include virtual="/manager/inc/inc.ui.head.asp"-->
</head>
<body class="skin-2">	
	<div id="wrapper" class="white-bg">
		
		<div id="page-wrapper" class="white-bg">

			<div class="table-responsive">
				<form role="form" id="frm_adminList_ifm" name="frm_adminList_ifm" method="post" action="UserAuth_ifm_Proc.asp">
					<input type="hidden" id="CD_ADMINID" name="CD_ADMINID" value="<%=CD_ADMINID%>">
					<input type="hidden" id="MODE" name="MODE">
					<input type="hidden" id="RTN_STR" name="RTN_STR" value="UserAuth.asp?CD_ADMINID=<%=CD_ADMINID%>&NM_NAME=<%=NM_NAME%>">
					<input type="hidden" id="nm_string" name="nm_string">
					
					<table id="table_check" name="table_check" class="table">
						<tr>
							<td colspan="5" class="text-left">
							<%If CD_ADMINID <> "" Then%>
								[<span class="text-success"><strong><%=NM_NAME%></strong></span>] 님의 사용 중인 메뉴 리스트
							<%Else%>
								<span class="text-danger"><strong>* 사용자를 선택해 주세요.</strong></span>
							<%End If%>
							</td>
						</tr>
						<tr>
							<td colspan="5" class="text-left">
								<button type="button" class="btn btn-success btn-sm" onclick="fn_checkAll(true);"><i class="fa fa-check-square-o"></i> 전체선택</button>
								<button type="button" class="btn btn-default btn-sm" onclick="fn_checkAll(false);"><i class="fa fa-square"></i> 선택해제</button>
								<button type="button" class="btn btn-primary btn-sm" onclick="fn_Edit('I','');"><i class="fa fa-check"></i> 저장</button>
							</td>
						</tr>
						<tr>
							<td class="text-center"><strong>선택</strong></td>
							<!--<td align="center">메뉴 구분</td>-->
							<td class="text-center"><strong>메뉴명</strong></td>
						</tr>
				<%
				If IsArray(ArrMenuList) Then
					For iloop = 0 To UBound(ArrMenuList,2) 
						CD_FORM		= ArrMenuList(0,iloop)
						CD_PFORM	= ArrMenuList(1,iloop)
						CD_ROOT		= ArrMenuList(2,iloop)
						NO_DEPTH	= ArrMenuList(3,iloop)
						FORMPART	= ArrMenuList(4,iloop)		
						NM_FORM		= ArrMenuList(5,iloop)
						YN_USEDMENU	= ArrMenuList(13,iloop)
				%>
						<!--<tr bgColor="<%=getBoolean(NO_DEPTH="0","#76A4C0",getBoolean(NO_DEPTH="1","#e7eff4","#ffffff"))%>">-->
						<tr class="<%=getBoolean(NO_DEPTH="0","gray-bg text-success",getBoolean(NO_DEPTH="1","white-bg text-info","white-bg"))%>">
							<td class="text-center">
								<input type="checkbox" id="<%=CD_FORM%>" name="<%=CD_FORM%>" value="<%=CD_FORM&"@"&CD_PFORM&"@"&CD_ROOT%>" <%=getBoolean(YN_USEDMENU="Y","checked","")%> OnClick="fn_autoCheck('<%=CD_FORM%>','<%=CD_PFORM%>','<%=CD_ROOT%>','<%=NO_DEPTH%>');">
							</td>
							<!--<td align="center"><%'=FORMPART%></td>-->
							<td class="text-left">
								<%If NO_DEPTH = "0" Then %>
									<%=NM_FORM%>
								<%ElseIf NO_DEPTH = "1" Then %>
									<%="&nbsp; &nbsp; &nbsp;" & NM_FORM%>
								<%ElseIf NO_DEPTH = "2" Then %>
									<%="&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" & NM_FORM%>
								<%End If %>
								
							</td>
						</tr>
				<%
					Next
				End If
				%>			

						<tr>
							<td colspan="5" class="text-left">
								<button type="button" class="btn btn-success btn-sm" onclick="fn_checkAll(true);"><i class="fa fa-check-square-o"></i> 전체선택</button>
								<button type="button" class="btn btn-default btn-sm" onclick="fn_checkAll(false);"><i class="fa fa-square"></i> 선택해제</button>
								<button type="button" class="btn btn-primary btn-sm" onclick="fn_Edit('I','');"><i class="fa fa-check"></i> 저장</button>
							</td>
						</tr>
					</table>
				</form>
			</div>
		</div>

	<script type="text/javascript">
		//# 메뉴 리스트 저장
		function fn_Edit(MODE, INDEX)
		{	
			document.getElementById("MODE").value = MODE;
			document.frm_adminList_ifm.submit();
		}
		
		//# 카테고리별 조회
		function fn_GUBUN(CD_FORM)
		{
			document.location.href="adminList_ifm.asp?CD_FORM=" + CD_FORM;
		}

		//# 리스트 전체 선택 / 선택 해제
		function fn_checkAll(flag) {
			$("input[type='checkbox']").prop("checked", flag);
		}
		
		//# 하위메뉴 체크 시 상위 메뉴 자동 체크
		function fn_autoCheck(CD_FORM, CD_PFORM, CD_ROOT, NO_DEPTH)
		{
			var form = document.frm_adminList_ifm;
			//var CHECK_MENU = form.tags("INPUT");
			var CHECK_MENU = $("input[type='checkbox']", form);

			var arrLevel;

			switch (NO_DEPTH) {
				case "0":
					for (var i = 0 ; i < CHECK_MENU.length ; i++ ) {
						arrLevel = CHECK_MENU[i].value.split("@");

						// 하위 메뉴 체크
						if (arrLevel[1] == CD_FORM || arrLevel[2] == CD_FORM) {
							CHECK_MENU[i].checked = true;
						}
						// 하위 메뉴 체크 해제
						if (document.getElementById(CD_FORM).checked == false && (arrLevel[1] == CD_FORM || arrLevel[2] == CD_FORM)) {
							CHECK_MENU[i].checked = false;
						}
					}					
				break;
				
				case "1":
					for (var i = 0 ; i < CHECK_MENU.length ; i++ ) {
						arrLevel = CHECK_MENU[i].value.split("@");

						// 상위 메뉴 체크
						if (arrLevel[0] == CD_PFORM) {
							CHECK_MENU[i].checked = true;
						}
						// 하위 메뉴 체크
						if (arrLevel[1] == CD_FORM) {
							CHECK_MENU[i].checked = true;
						}
						// 하위 메뉴 체크 해제
						if (document.getElementById(CD_FORM).checked == false && arrLevel[1] == CD_FORM) {
							CHECK_MENU[i].checked = false;
						}
					}
				break;
				
				case "2":
					for (var i = 0 ; i < CHECK_MENU.length ; i++ ) {
						arrLevel = CHECK_MENU[i].value.split("@");

						// 상위 메뉴 체크
						if (arrLevel[0] == CD_ROOT || arrLevel[0] == CD_PFORM) {
							CHECK_MENU[i].checked = true;				
						}
					}
				break;
			}	
		}
	</script>
</body>
</html>		
<%
Set objADO = Nothing
%>