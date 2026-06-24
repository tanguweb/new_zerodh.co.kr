<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_CodeInfo.asp" -->
<!--#include virtual="/manager/inc/inc.ui.head.asp"-->
<%
	'########################################################################################
	'#	File		: /manager/auth/adminmenu.asp
	'#  Create		: 조영준 / 2011.01.07
	'#	Info		: 어드민 메뉴 등록
	'#	Update		:
	'#	Update Memo	:
	'########################################################################################

	If m_UserID <> "sysadmin" Then
		Response.Write "<script language='JavaScript'>"
		Response.Write "	alert('죄송합니다. \n\n시스템 관리자만 사용 가능 합니다.');history.back();"
		Response.Write "</script>"
		Response.End
	End If

	Dim objADO, Rs, SQL, SQL1, SQL2, iloop, jloop
	Dim ArrMenuList, ArrFormPart, ArrMainMenu, ArrSubMenu
	Dim CD_FORM, NO_DEPTH, CD_PFORM, CD_PART, CD_NAMESPACE, CD_EDITION, NM_FORM, NM_URL, NO_ORDER, YN_USER, YN_USE, NM_ICON, CD_TARGET
	Dim MAX_DEPTH : MAX_DEPTH = 2 '// 0, 1, 2

	Set objADO = new clsADO

	objADO.setConString(m_DB)

	'// 메뉴 전체 리스트
	SQL = ""
	SQL = SQL & " 		;WITH MENU_ADMIN(CD_FORM,CD_PFORM,NO_DEPTH,CD_PART,NM_FORM,NM_URL,CD_NAMESPACE,CD_EDITION,CD_TARGET,NO_ORDER,MM_HELP,YN_USER,YN_USE,levels,sort) "
	SQL = SQL & " 		AS "
	SQL = SQL & " 		("
	SQL = SQL & " 			SELECT "
	SQL = SQL & " 				CD_FORM,CD_PFORM,NO_DEPTH,CD_PART"
	SQL = SQL & " 				,NM_FORM,NM_URL,CD_NAMESPACE,CD_EDITION"
	SQL = SQL & " 				,CD_TARGET,NO_ORDER,MM_HELP,YN_USER,YN_USE"
	SQL = SQL & " 				,0 AS levels"
	SQL = SQL & " 				,CONVERT(VARCHAR(255), CD_FORM) AS [sort]"
	SQL = SQL & " 			FROM T_MENU"
	SQL = SQL & " 			WHERE NO_DEPTH = '0'"

	SQL = SQL & " 			UNION ALL"

	SQL = SQL & " 			SELECT "
	SQL = SQL & " 				A.CD_FORM,A.CD_PFORM,A.NO_DEPTH,A.CD_PART"
	SQL = SQL & " 				,A.NM_FORM,A.NM_URL,A.CD_NAMESPACE,A.CD_EDITION"
	SQL = SQL & " 				,A.CD_TARGET,A.NO_ORDER,A.MM_HELP,A.YN_USER,A.YN_USE"
	SQL = SQL & " 				,(levels + 1)"
	SQL = SQL & " 				,CONVERT(VARCHAR(255), B.sort + ' ' + A.CD_FORM) AS [sort]"
	SQL = SQL & " 			FROM T_MENU AS A"
	SQL = SQL & " 			INNER JOIN MENU_ADMIN AS B ON B.CD_FORM = A.CD_PFORM"
	SQL = SQL & " 		)"
	SQL = SQL & " 		SELECT "
	SQL = SQL & " 			CD_FORM,CD_PFORM,NO_DEPTH,CD_PART"
	SQL = SQL & " 			,NM_FORM,NM_URL,CD_NAMESPACE,CD_EDITION"
	SQL = SQL & " 			,CD_TARGET,NO_ORDER,MM_HELP,YN_USER,YN_USE "
	SQL = SQL & " 		FROM MENU_ADMIN "
	SQL = SQL & " 		ORDER BY NO_ORDER ASC, sort ASC "

	objADO.setSql(SQL)
	ArrMenuList = objADO.getArrRs()

	'// 상위 메뉴 리스트(메인)
	SQL1 = SQL1 & " SELECT CD_FORM, CD_PFORM, NO_DEPTH, CD_PART, NM_FORM "
	SQL1 = SQL1 & " FROM T_MENU										 "
	SQL1 = SQL1 & " WHERE YN_USE = 'Y' AND NO_DEPTH = '0'				 "

	objADO.setSql(SQL1)
	ArrMainMenu = objADO.getArrRs()

	'// 상위 메뉴 리스트(서브)
	SQL2 = SQL2 & " SELECT CD_FORM, CD_PFORM, NO_DEPTH, CD_PART, NM_FORM "
	SQL2 = SQL2 & " FROM T_MENU										 "
	SQL2 = SQL2 & " WHERE YN_USE = 'Y' AND NO_DEPTH = '1'				 "

	objADO.setSql(SQL2)
	ArrSubMenu = objADO.getArrRs()
%>
	<script language="Javascript" src="/js/HttpRequest.js"></script>
	<script language="javascript" src="/js/jsxml.js"></script>
    <script language="javascript" src="/js/json.js"></script>
	<Script Language = "JavaScript">
		//# Depth별 상위메뉴 초기화 및 리스트 가져오기
		function fn_changeParent(no)
		{
			var NO_DEPTH = document.getElementById("NO_DEPTH_" + no);

			if (NO_DEPTH.value == "0") {
				//document.getElementById("CD_PFORM_" + no).value = "";
				var form = document.frm_Menu
				var NO_DEPTH = eval("form.NO_DEPTH_" + no).value;
				var CD_PFORM = eval("form.CD_PFORM_" + no);

				//상위메뉴 초기화
				for (var i = CD_PFORM.length; i > 0; i--) {
					CD_PFORM.options[i] = null;
				}
				CD_PFORM.selectIndex = 0;

				return;
			}
			else
			{
				var form = document.frm_Menu
				var NO_DEPTH = eval("form.NO_DEPTH_" + no).value;
				var CD_PFORM = eval("form.CD_PFORM_" + no);

				//상위메뉴 초기화
				for (var i = CD_PFORM.length; i > 0; i--) {
					CD_PFORM.options[i] = null;
				}
				CD_PFORM.selectIndex = 0;

				try {
					var result = "";
					hr = new HttpRequest2(); // Version 1.1.0701241152
					hr.url = "/manager/auth/XML_ParentCode.asp";	// XML 생성 파일
					hr.params = "NO_DEPTH=" + NO_DEPTH;
					hr.method = "post";
					result = hr.response();
					result = result.replace(/<[\\s]*(\/)*(script)[\\s]*([^>]*)[\\s]*>/gi,"&lt;$1$2$3&gt;");
					//alert(result);

					var _xmlDoc = new REXML(result); 	//결과를 XML로 변환작업
					var _element = _xmlDoc.rootElement; 	//하위 Element를 가져옴

					var CD_PFORMCnt = _element.childElements[0].childElements.length;

					for (var i = 0, j = CD_PFORMCnt; i < j; i++) {
						var _items = _element.childElements[0].childElements[i];
						var CD_PFORM_Info = new Option();
						CD_PFORM_Info.value = _items.childElements[0].getText();
						CD_PFORM_Info.text = _items.childElements[1].getText();
						CD_PFORM.options.add(CD_PFORM_Info);
					}
				}
				catch (e) {
					alert(e.name + " : " + e.message);
				}
			}
		}

		//# 저장, 수정, 삭제하기
		function f_Edit(MODE, CD_FORM)
		{
			var CD_FORM;
			var NO_DEPTH;
			var CD_PFORM;
			var NM_FORM;
			var NM_URL;
			var NO_ORDER;

			if (MODE=="I")
			{
				CD_FORM	  = document.getElementById("CD_FORM_0");
				NO_DEPTH  = document.getElementById("NO_DEPTH_0");
				CD_PFORM  = document.getElementById("CD_PFORM_0");
				//CD_PART	  = document.getElementById("CD_PART_0");
				NM_FORM   = document.getElementById("NM_FORM_0");
				NM_URL	  = document.getElementById("NM_URL_0");
				NO_ORDER  = document.getElementById("NO_ORDER_0");

				if (CD_FORM == "")
				{
					alert("메뉴코드를 입력해주세요.");
					CD_FORM.focus();
					return;
				}

				if (ValNumber(NO_DEPTH,"뎁스를",true) == false) { return; }

				/*if (CD_PFORM.value == "")
				{
					alert("상위메뉴를 선택해주세요.");
					CD_PFORM.focus();
					return;
				}*/

				/*if (CD_PART.value == "")
				{
					alert("메뉴 파트를 선택해주세요.");
					CD_PART.focus();
					return;
				}*/

				if (NM_FORM.value == "")
				{
					alert("메뉴명을 입력하세요.");
					NM_FORM.focus();
					return;
				}

				if (NO_DEPTH.value == "0" || NO_DEPTH.value == "1")
				{
					NM_URL.value = "";
				}
				else
				{
					if (NM_URL.value == "") { alert("링크를 입력하세요."); NM_URL.focus(); return; }
				}

				if (ValNumber(NO_ORDER, "정렬순서를", true)== false) { return; };

				document.getElementById("CD_FORM").value = document.getElementById("CD_FORM_0").value;
				document.getElementById("nm_string").value = "0";
			}

			else if(MODE=="U")
			{
				NO_DEPTH  = document.getElementById("NO_DEPTH_" + CD_FORM);
				CD_PFORM  = document.getElementById("CD_PFORM_" + CD_FORM);
				NM_FORM   = document.getElementById("NM_FORM_" + CD_FORM);
				NM_URL	  = document.getElementById("NM_URL_" + CD_FORM);
				NO_ORDER  = document.getElementById("NO_ORDER_" + CD_FORM);

				//CD_NAMESPACE	  = document.getElementById("CD_NAMESPACE_" + CD_FORM);
				//alert(CD_FORM);
				//alert(CD_NAMESPACE.value);
				//return;
				//alert(ValNumber(NO_DEPTH,"뎁스",true));
				if (ValNumber(NO_DEPTH,"뎁스",true) == false) { return; }
				//if (CD_PFORM.value != "" && ValNumber(CD_PFORM,"상위메뉴코드를",true) == false) { return; }
				if (NM_FORM.value == "") { alert("메뉴명을 입력하세요."); NM_FORM.focus(); return; }

				//if (NO_DEPTH.value == "0" || NO_DEPTH.value == "1")
				//{
				//	NM_URL.value = "";
				//}
				//else
				//{
				//	if (NM_URL.value == "") { alert("링크를 입력하세요."); NM_URL.focus(); return; }
				//}

				if (ValNumber(NO_ORDER, "정렬순서를", true)== false) { return; };

				document.getElementById("CD_FORM").value = CD_FORM;
				document.getElementById("nm_string").value = CD_FORM;
			}
			else if(MODE == "D")
			{
				if (confirm("삭제 하시겠습니까?") == true)
				{
					document.getElementById("nm_string").value = CD_FORM;
				}
				else
				{
					return;
				}
			}

			document.getElementById("MODE").value = MODE;

			//document.frm_Menu.target = "hd_frame";
			document.frm_Menu.method = "post";
			document.frm_Menu.action = "UserMenu_proc.asp";
			document.frm_Menu.submit();
		}
	</Script>
</HEAD>
<BODY>
<form id="frm_Menu" name="frm_Menu">
	<input type="hidden" id="MODE" name="MODE">
	<input type="hidden" id="RTN_STR" name="RTN_STR" value="UserMenu.asp">
	<input type="hidden" id="CD_FORM" name="CD_FORM">
	<input type="hidden" id="nm_string" name="nm_string">

	<font color="red">
		※ DEPTH 가 0 이면 상위 메뉴 선택 없음 <br>
		※ DEPTH 가 2 미만이면 URL 입력 필요없음 <br>
		※ 사용자 코드여부 (예 : 메뉴권한 관리에서 조회됨 / 아니오 : 메뉴권한 관리에 안보임) <br>
		※ 사용여부 (아니오 : 좌측 메뉴에 보이지 않음) <br>
		※ 정렬순서 중요!!!! <br>
	</font>

	<table cellSpacing="0" cellPadding="1" align="left" class="table_head">
	<tr class="table_head_title">
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>메뉴코드</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>DEPTH</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>상위메뉴</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>메뉴이름</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>URL</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>TARGET</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>NAMESPACE</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>EDITION</b></td>
		<!--<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>ICON</b></td>-->
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>정렬순서</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>사용자 코드여부</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>사용여부</b></td>
		<td align="center" width="" style="border-bottom: 1px solid #ccc;"><b>저장</b></td>
	</tr>
	<tr class="table_head_title" onMouseOver="javascript:this.style.backgroundColor='#D4D4D4'" onMouseOut="javascript:this.style.backgroundColor=''">
		<td align="center"><input type="text" class="input" id="CD_FORM_0" name="CD_FORM_0" size="20" style="ime-mode:inactive"></td>
		<td align="center">
			<select id="NO_DEPTH_0" name="NO_DEPTH_0" class="select" OnChange="fn_changeParent('0');">
				<option value="">선택</option>
			<%For jloop = 0 To MAX_DEPTH %>
				<option value="<%=jloop%>"><%=jloop%></option>
			<%Next %>
			</select>
		</td>
		<td align="center">
			<select id="CD_PFORM_0" name="CD_PFORM_0" class="select" style="width:85px;">
				<option value="">선택</option>
			</select>
		</td>
		<td align="center"><input type="text" class="input" id="NM_FORM_0" name="NM_FORM_0" size="20" style="ime-mode:active"></td>
		<td align="center"><input type="text" class="input" id="NM_URL_0" name="NM_URL_0" size="50" style="ime-mode:inactive"></td>
		<td align="center"><input type="text" class="input" id="CD_TARGET_0" name="CD_TARGET_0" size="15" style="ime-mode:inactive"></td>
		<td align="center"><input type="text" class="input" id="CD_NAMESPACE_0" name="CD_NAMESPACE_0" size="15" style="ime-mode:inactive"></td>
		<td align="center"><input type="text" class="input" id="CD_EDITION_0" name="CD_EDITION_0" size="15" style="ime-mode:inactive"></td>
		<!--<td align="center"><input type="text" class="input" id="NM_ICON_0" name="NM_ICON_0" size="15" style="ime-mode:inactive"></td>-->
		<td align="center"><input type="text" class="input" id="NO_ORDER_0" name="NO_ORDER_0" value="1000" size="6"></td>
		<td align="center">
			<select name="YN_USER_0" class="select" >
				<option value="Y" selected>예</option>
				<option value="N">아니오</option>
			</select>
		</td>
		<td align="center">
			<select name="YN_USE_0" class="select" >
				<option value="Y" selected>예</option>
				<option value="N">아니오</option>
			</select>
		</td>
		<td align="center"><input type="button" class="button" value = "추가" onClick="f_Edit('I','0');"></td>
	</tr>
	<tr class="table_head_title">
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>메뉴코드</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>DEPTH</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>상위메뉴</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>메뉴이름</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>URL</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>TARGET</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>NAMESPACE</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>EDITION</b></td>
		<!--<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>ICON</b></td>-->
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>정렬순서</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>사용자 코드여부</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>사용여부</b></td>
		<td align="center" width="" style="border-top: 2px solid #ccc; border-bottom: 1px solid #ccc;"><b>관리</b></td>
	</tr>
<%
	If IsArray(ArrMenuList) Then
		For iloop = 0 To UBound(ArrMenuList,2)
			CD_FORM			= ArrMenuList(0,iloop)
			CD_PFORM		= ArrMenuList(1,iloop)
			NO_DEPTH		= ArrMenuList(2,iloop)
			CD_PART			= ArrMenuList(3,iloop)
			NM_FORM			= ArrMenuList(4,iloop)
			NM_URL			= ArrMenuList(5,iloop)
			CD_NAMESPACE	= ArrMenuList(6,iloop)
			CD_EDITION		= ArrMenuList(7,iloop)
			'NM_ICON			= ArrMenuList(8,iloop)
			CD_TARGET       = ArrMenuList(8,iloop)
			NO_ORDER		= ArrMenuList(9,iloop)
			YN_USER			= ArrMenuList(11,iloop)
			YN_USE			= ArrMenuList(12,iloop)

%>
	<tr class="table_head_title" style="background-color:<%=getBoolean(NO_DEPTH="0","#76A4C0",getBoolean(NO_DEPTH="1","#e7eff4","#ffffff"))%>">
		<td align="left">
			<input type="hidden" id="CD_FORM_<%=CD_FORM%>" name="CD_FORM_<%=CD_FORM%>" value="<%=CD_FORM%>" size="6"><%=CD_FORM%>
		</td>
		<td align="center">
			<select id="NO_DEPTH_<%=CD_FORM%>" name="NO_DEPTH_<%=CD_FORM%>" class="select" OnChange="fn_changeParent('<%=CD_FORM%>');">
				<option value="">선택</option>
			<%For jloop = 0 To MAX_DEPTH %>
				<option value="<%=jloop%>" <%=getBoolean(NO_DEPTH = jloop,"selected","")%>><%=jloop%></option>
			<%Next %>
			</select>
		</td>
		<td align="center">
			<select id="CD_PFORM_<%=CD_FORM%>" name="CD_PFORM_<%=CD_FORM%>" class="select" style="width:85px;">
				<option value="">선택</option>
			<%
			If NO_DEPTH = "1" And IsArray(ArrMainMenu) Then
				For jloop = 0 To UBound(ArrMainMenu,2)
			%>
				<option value="<%=ArrMainMenu(0,jloop)%>" <%=getBoolean(ArrMainMenu(0,jloop)=CD_PFORM,"selected","")%>><%=ArrMainMenu(4,jloop)%></option>
			<%
				Next
			ElseIf NO_DEPTH = "2" And IsArray(ArrSubMenu) Then
				For jloop = 0 To UBound(ArrSubMenu,2)
			%>
				<option value="<%=ArrSubMenu(0,jloop)%>" <%=getBoolean(ArrSubMenu(0,jloop)=CD_PFORM,"selected","")%>><%=ArrSubMenu(4,jloop)%></option>
			<%
				Next
			End If
			%>
			</select>
		</td>
		<td align="center"><input type="text" class="input" id="NM_FORM_<%=CD_FORM%>" name="NM_FORM_<%=CD_FORM%>" value="<%=NM_FORM%>" size="20" style="ime-mode:active"></td>
		<td align="center"><input type="text" class="input" id="NM_URL_<%=CD_FORM%>" name="NM_URL_<%=CD_FORM%>" value="<%=NM_URL%>" size="50" style="ime-mode:inactive"></td>
		<td align="center"><input type="text" class="input" id="CD_TARGET_<%=CD_FORM%>" name="CD_TARGET_<%=CD_FORM%>" value="<%=CD_TARGET%>" size="15" style="ime-mode:inactive"></td>
		<td align="center"><input type="text" class="input" id="CD_NAMESPACE_<%=CD_FORM%>" name="CD_NAMESPACE_<%=CD_FORM%>" value="<%=CD_NAMESPACE%>" size="15" style="ime-mode:inactive"></td>
		<td align="center"><input type="text" class="input" id="CD_EDITION_<%=CD_FORM%>" name="CD_EDITION_<%=CD_FORM%>" value="<%=CD_EDITION%>" size="45" style="ime-mode:inactive"></td>
		<td align="center"><input type="text" class="input" id="NO_ORDER_<%=CD_FORM%>" name="NO_ORDER_<%=CD_FORM%>" value="<%=NO_ORDER%>" size="6"></td>
		<td align="center">
			<select name="YN_USER_<%=CD_FORM%>" class="select" >
				<option value="Y" <%=getBoolean(YN_USER="Y","selected","")%>>예</option>
				<option value="N" <%=getBoolean(YN_USER="N","selected","")%>>아니오</option>
			</select>
		</td>
		<td align="center">
			<select name="YN_USE_<%=CD_FORM%>" class="select" >
				<option value="Y" <%=getBoolean(YN_USE="Y","selected","")%>>예</option>
				<option value="N" <%=getBoolean(YN_USE="N","selected","")%>>아니오</option>
			</select>
		</td>
		<td align="center"><input type="button" class="button" value="수정" onClick="f_Edit('U','<%=CD_FORM%>');"></td>
	</tr>
<%
		Next
	End If
%>
	</table>
</form>
 </BODY>
</HTML>
<iframe type="hiddenframe" id="hd_frame" name="hd_frame" style="display:none;"></iframe>
<%
Set objADO = Nothing
%>
