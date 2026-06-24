<!-- #include virtual = "/Include/Config.asp" -->
<!-- #include virtual = "/function/fn_AdminLoginCheck.asp" -->
<!-- #include virtual = "/function/fn_PageNavigation.asp"-->
<%
	'########################################################################################
	'#	File		: /manager/auth/adminList.asp
	'#  Create		: domoyosi / 2015.12
	'#	Info		: 어드민 메뉴리스트
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################
	Dim objADO, ArrAdminMember, ArrAdminMenu, SQL, iloop, jloop
	Dim NO_CNT, CD_ADMINID, NM_NAME, NO_TEL, NM_EMAIL, YN_LOGIN, YN_USE
	
	NM_NAME =	ReqQ("NM_NAME")

	Dim page, startpage, tn, pgcount, nextpage, prevpage
	Const pgsize = 15	'출력 라인 수
	Const setsize = 10	'페이징 수

	If IsNumeric(ReqQ("page")) = False then  
		page = 1  
		startpage = 1
	Else 
		page = CInt(ReqQ("page"))
		startpage = Int(page/setsize)
			If startpage = (page/setsize) Then 
				startpage = page-setsize + 1
			Else 
				startpage = Int(page/setsize)*setsize + 1
			End If 
	End If

	Set objADO = new clsADO
	
	SQL = ""
	SQL = SQL & " DECLARE @CNT INT											"
	SQL = SQL & " DECLARE @SNUM INT											"
	SQL = SQL & " DECLARE @ENUM INT											"
	SQL = SQL & "															"
	SQL = SQL & " SELECT @CNT = COUNT(*)									"
	SQL = SQL & " FROM T_ADMIN												"
	SQL = SQL & " WHERE YN_LOGIN = 'Y' AND YN_USE = 'Y'						"
	SQL = SQL & "															"
	SQL = SQL & " SET @SNUM = @CNT - ( " & pgsize & " * " & page & " ) + 1	"
	SQL = SQL & " SET @ENUM = @SNUM + " & pgsize & " - 1					"
	SQL = SQL & "															"
	SQL = SQL & " SELECT 													"
	SQL = SQL & " 	@CNT AS CNT, *											"
	SQL = SQL & " FROM														"
	SQL = SQL & " (															"
	SQL = SQL & " SELECT													"
	SQL = SQL & " 	ROW_NUMBER() OVER (ORDER BY CD_ADMINID DESC) AS [No]	" '1
	SQL = SQL & " 	, CD_ADMINID, NM_ADMINPW, NM_NAME, NO_TEL, NM_EMAIL		" '6
	SQL = SQL & "	, YN_LOGIN, YN_USE										" '8
	SQL = SQL & " FROM T_ADMIN												"
	SQL = SQL & " WHERE YN_LOGIN = 'Y' AND YN_USE = 'Y'						"
	SQL = SQL & " ) AS L													"
	SQL = SQL & " WHERE No BETWEEN @SNUM AND @ENUM							"
	SQL = SQL & " ORDER BY													"
	SQL = SQL & " 	No DESC													"

	objADO.setSql(SQL)
	ArrAdminMember = objADO.getArrRs()

	If IsArray(ArrAdminMember) Then
		tn =  ArrAdminMember(0,0)							'총 페이지수
		pgcount = - INT( - ( tn/pgsize ) )					'페이지수 계산
		If (page > pgcount) Then page = pgcount End If 		'페이지수 이상의 페이지를 요구하면 마지막 페이지로

		'이전page 다음page
		prevPage = page - ((page-1) Mod setsize) - setsize
		If prevPage < 1 Then
			prevPage = 0
		End IF
		nextPage = page - ((page-1) Mod setsize) + setsize
		If nextPage > pgcount Then
			nextPage = 0
		End If
	Else
		tn = 0
		pgcount = 0
	End If
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
                    <h2>사용자별 메뉴 권한관리</h2>
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
                            <strong>사용자별 메뉴 권한관리</strong>
                        </li>
                    </ol>
                </div>
            </div>
			

			<div class="wrapper wrapper-content animated fadeInRight">
				<div class="row">
					<div class="col-lg-6">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-list"></i> 사용자 리스트
									<small>
									<%
									If IsArray(ArrAdminMember) Then
										Response.Write " (전체 관리자 수 : <span class='text-info'><strong>" & tn & "</strong></span> 명)"
									End If 
									%>
									</small>
								</h5>
							</div>
							<div class="ibox-content">

								<div class="table-responsive">
									<form id="frm_adminList" name="frm_adminList">
										<input type="hidden" id="MODE" name="MODE">
										<input type="hidden" id="RTN_STR" name="RTN_STR" value="UserAuth.asp?CD_ADMINID=<%=CD_ADMINID%>">
										<input type="hidden" id="nm_string" name="nm_string">
										<table class="table table-hover table-striped">
											<thead>
												<tr>
													<th class="text-center">No.</th>
													<th class="text-center">ID</th>
													<th class="text-center">이름</th>
													<!--<th class="text-center">연락처</th>-->
													<!--<th class="text-center">이메일</th>-->
													<th class="text-center">어드민로그인여부</th>
													<th class="text-center">사용여부</th>
												</tr>
											</thead>
											<%	
												If IsArray(ArrAdminMember) Then
													For iloop = 0 To UBound(ArrAdminMember,2) 
														NO_CNT		= ArrAdminMember(1,iloop)
														CD_ADMINID	= ArrAdminMember(2,iloop)
														NM_NAME		= ArrAdminMember(4,iloop)
														NO_TEL		= ArrAdminMember(5,iloop)
														NM_EMAIL	= ArrAdminMember(6,iloop)
														YN_LOGIN	= ArrAdminMember(7,iloop)
														YN_USE		= ArrAdminMember(8,iloop)
											%>
											<tr id="row_<%=iloop%>" class="<%=getBoolean(ReqQ("CD_ADMINID")=CD_ADMINID,"text-success","")%>" style="">
												<input type="hidden" id="CD_ADMINID_<%=CD_ADMINID%>" name="CD_ADMINID_<%=CD_ADMINID%>" value="<%=CD_ADMINID%>">
												<td class="text-center"><%=NO_CNT%></td>
												<td class="text-center" OnClick="fn_adminDetail('<%=CD_ADMINID%>','<%=NM_NAME%>','<%=iloop%>');" style="cursor:pointer;"><%=CD_ADMINID%></td>
												<td class="text-center"><%=NM_NAME%></td>
												<!--
												<td><%=NO_TEL%></td>
												<td><%=NM_EMAIL%></td>
												-->
												<td class="text-center"><%=getboolean(YN_LOGIN="Y","예","아니오")%></td>
												<td class="text-center"><%=getboolean(YN_USE="Y","예","아니오")%></td>
											</tr>

										<%	
												Next
											End If
										%>
											<tr>
												<td colspan="7">
												<%
													Call PageNavigation("/manager/auth/UserAuth.asp","", tn, pgcount, startpage, page, setsize, prevpage, nextpage)
												%>
												</td>
											</tr>

										</table>
									</form>
								</div>	

							</div>
						</div>
					</div>


					<div class="col-lg-6">
						
						<div class="ibox float-e-margins">
							<div class="ibox-title">
								<h5><i class="fa fa-list"></i> 사용자별 메뉴 권한관리 <small>사용자별로 메뉴의 사용여부를 관리합니다.</small></h5>
							</div>
							<div class="ibox-content">
								<iframe id="ifm_adminMenu" name="ifm_adminMenu" frameborder="0" marginheight="0" marginwidth="0" scrolling="yes" style="width:100%; height:750px; border:1px solid #ccc;" src="UserAuth_ifm.asp?CD_ADMINID=<%=getBoolean(ReqQ("CD_ADMINID")<>"",ReqQ("CD_ADMINID"),"")%>&NM_NAME=<%=getBoolean(ReqQ("NM_NAME")<>"",ReqQ("NM_NAME"),"")%>"></iframe>
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
	//# 사용자별 메뉴 사용 권한 iframe으로 보여주기
		function fn_adminDetail(CD_ADMINID,NM_NAME,index)
		{	
			<%for iloop = 0 to UBound(ArrAdminMember,2)%>
				//document.getElementById("row_" + <%=iloop%>).style.fontWeight = "";
				$("#row_" + <%=iloop%>).removeClass("text-success");
			<%next%>
			//document.getElementById("row_" + index).style.fontWeight = "bold";
			$("#row_" + index).addClass("text-success");
			//document.getElementById("ifm_adminMenu").src = "UserAuth_ifm.asp?CD_ADMINID=" + CD_ADMINID + "&NM_NAME=" + escape(NM_NAME);
			$("#ifm_adminMenu").attr("src", "UserAuth_ifm.asp?CD_ADMINID=" + CD_ADMINID + "&NM_NAME=" + escape(NM_NAME));
		}
	</script>
</body>
</html>		
<%
Set objADO = Nothing
%>